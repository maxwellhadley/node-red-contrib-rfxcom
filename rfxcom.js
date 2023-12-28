/*
    Copyright (c) 2014 .. 2019, Maxwell Hadley
    All rights reserved.

    Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following
    conditions are met:

    1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following
    disclaimer.

    2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following
    disclaimer in the documentation and/or other materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
    IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
    FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
    CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
    DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
    DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
    CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
    USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

module.exports = function (RED) {
    "use strict";
    const rfxcom = require("rfxcom");

// Set the rfxcom debug option from the environment variable
    let enableAllDebug = false;
    if (process.env.hasOwnProperty("RED_DEBUG") && process.env.RED_DEBUG.indexOf("rfxcom") >= 0) {
        enableAllDebug = true;
    }

// The config node holding the (serial) port device path for one or more rfxcom family nodes
    function RfxtrxPortNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.rfyVenetianMode = n.rfyVenetianMode || "EU";
        this.enableDebug = n.enableDebug || false;
    }

// Register the config node
    RED.nodes.registerType("rfxtrx-port", RfxtrxPortNode);

// An object maintaining a pool of config nodes
    const rfxcomPool = function () {
        let pool = {};

        const connectTo = function (rfxtrx, node) {
            //noinspection JSUnusedLocalSymbols
            rfxtrx.initialise(function (error, response, sequenceNumber) {
                node.log("connected: Serial port " + rfxtrx.device);
                if (pool[rfxtrx.device].intervalTimer !== null) {
                    clearInterval(pool[rfxtrx.device].intervalTimer);
                    pool[rfxtrx.device].intervalTimer = null;
                }
            });
        };

        return {
            get: function (node, rfxtrxPort) {
                // Returns the RfxCom object associated with port, or creates a new RfxCom object,
                // associates it with the port, and returns it. 'rfxtrxPort' is the config node holding
                // the device file path to the pseudo-serialport, e.g. '/dev/tty.usb-123456'
                let rfxtrx;
                const port = rfxtrxPort.port;
                if (!pool[port]) {
                    rfxtrx = new rfxcom.RfxCom(port, (enableAllDebug || rfxtrxPort.enableDebug) ? {debug: true} : {});
                    rfxtrx.on("connecting", function () {
                        node.log("connecting to " + port);
                        pool[port].references.forEach(function (node) {
                            node.status({fill:"yellow",shape:"dot",text:"connecting..."});
                        });
                    });
                    rfxtrx.on("connectfailed", function (msg) {
                        if (pool[rfxtrx.device].intervalTimer === null) {
                            node.log("connect failed: " + msg);
                            pool[rfxtrx.device].intervalTimer = setInterval(function () {
                                connectTo(rfxtrx, node)
                            }, rfxtrx.initialiseWaitTime);
                        }
                    });
                    rfxtrx.on("status", function (status) {
                        rfxtrx.receiverType = status.receiverType;
                        rfxtrx.firmwareVersion = status.firmwareVersion;
                        rfxtrx.firmwareType = status.firmwareType;
                        rfxtrx.enabledProtocols = status.enabledProtocols;
                        pool[port].references.forEach(function (node) {
                                showConnectionStatus(node);
                            });
                    });
                    rfxtrx.on("response", function (message, seqnbr, responseCode) {
                        if (responseCode > 1) {
                            node.warn("RFXCOM: " + message + " (" + responseCode + ")");
                        }
                    });
                    rfxtrx.on("rfyremoteslist", function (list) {
                        if (list.length === 0) {
                            node.warn("RFXCOM: No RFY remotes are programmed in this device")
                        } else {
                            let message = "RFXCOM: RFY remotes in this device:";
                            list.forEach(function (entry) {
                                message = message + "\n  " + entry.remoteType + "/" + entry.deviceId;
                            });
                            node.warn(message);
                        }
                    });
                    rfxtrx.on("disconnect", function (msg) {
                        node.log("disconnected: " + msg);
                        pool[port].references.forEach(function (node) {
                                showConnectionStatus(node);
                            });
                        if (pool[rfxtrx.device].intervalTimer === null) {
                            pool[rfxtrx.device].intervalTimer = setInterval(function () {
                                connectTo(rfxtrx, node)
                            }, rfxtrx.initialiseWaitTime);
                        }
                    });
                    pool[port] = {rfxtrx: rfxtrx, references: [], intervalTimer: null};
                } else {
                    rfxtrx = pool[port].rfxtrx;
                }
                if (rfxtrx.connected === false && rfxtrx.initialising === false) {
                    connectTo(rfxtrx, node);
                }
                // Maintain a reference count for each RfxCom object
                pool[port].references.push(node);
                return pool[port].rfxtrx;
            },
            release: function (node, port) {
                // Decrement the reference count, and delete the RfxCom object if the count goes to 0
                if (pool[port]) {
                    pool[port].references.splice(pool[port].references.indexOf(node), 1);
                    if (pool[port].references.length <= 0) {
                        pool[port].rfxtrx.close();
                        pool[port].rfxtrx.removeAllListeners();
                        if (pool[port].intervalTimer !== null) {
                            clearInterval(pool[port].intervalTimer);
                            pool[port].intervalTimer = null;
                        }
                        delete pool[port].rfxtrx;
                        delete pool[port];
                    }
                }
            }
        }
    }();

    const releasePort = function (node) {
        // Decrement the reference count on the node port
        if (node.rfxtrxPort) {
            rfxcomPool.release(node, node.rfxtrxPort.port);
        }
    };

// Utility function: normalises the accepted representations of 'unit addresses' to be
// an integer, and ensures all variants of the 'group address' are converted to 0
    const parseUnitAddress = function (str) {
        if (str === undefined || /all|group|\+/i.test(str)) {
            return 0;
        } else {
            return Math.round(Number(str));
        }
    };

// This function takes a protocol name and returns the subtype number (defined by the RFXCOM
// API) for that protocol. It also creates the node-rfxcom object implementing the message packet type
// corresponding to that subtype (from the list provided), or re-uses a pre-existing object that implements it.
// These objects are stored in the transmitters property of rfxcomObject
    const getRfxcomSubtype = function (rfxcomObject, protocolName, transmitterPacketTypes, options) {
        let subtype = -1;
        if (rfxcomObject.transmitters.hasOwnProperty(protocolName) === false) {
            transmitterPacketTypes.forEach(function (packetType) {
                if (rfxcom[packetType][protocolName] !== undefined) {
                    subtype = rfxcom[packetType][protocolName];
                    rfxcomObject.transmitters[protocolName] = new rfxcom[packetType].transmitter(rfxcomObject, subtype, options);
                }
            });
        } else {
            subtype = rfxcomObject.transmitters[protocolName].subtype;
        }
        return subtype;
    };

// Convert a string containing a slash/delimited/path to an Array of (string) parts, removing any empty components
// Return value may be zero-length
    const stringToParts = function (str) {
        if (typeof str === "string") {
            return str.split('/').filter(function (part) {
                return part !== "";
            });
        } else {
            return [];
        }
    };

// Convert a string - the rawTopic - into a normalised form (an Array) so that checkTopic() can easily compare
// a topic against a pattern
    const normaliseTopic = function (rawTopic) {
        if (rawTopic === undefined || typeof rawTopic !== "string") {
            return [];
        }
        let parts = stringToParts(rawTopic);
        if (parts.length >= 1) {
            parts[0] = parts[0].trim().replace(/ +/g, '_').toUpperCase();
        }
        if (parts.length >= 2) {
            // handle houseCodes as a special case (X10, ARC, etc)
            if (/^[A-Z]$/i.test(parts[1])) {
                parts[1] = parseInt(parts[1].trim(), 36);
            } else {
                // ID is always in hexadecimal
                parts[1] = parseInt(parts[1].trim(), 16);
            }
        }
        if (parts.length >= 3) {
            if (/^0+$|all|group|^\+$/i.test(parts[2])) {
                return parts.slice(0, 2);
            }
            // handle Blyss groupcodes as a special case
            if (/^[A-Z]$/i.test(parts[2])) {
                parts[2] = parseInt(parts[2].trim(), 36);
            } else {
                // unitcodes always decimal
                parts[2] = parseInt(parts[2].trim(), 10);
            }
        }
        if (parts.length >= 4) {
            if (/^0+$|all|group|^\+$/i.test(parts[3])) {
                return parts.slice(0, 3);
            }
            // Blyss unitcodes always decimal
            parts[3] = parseInt(parts[3].trim(), 10);
        }
        // The return is always [ string, number, number, number ] - all parts optional
        return parts;
    };

// Normalise the supplied topic and check if it starts with the given pattern
    const normaliseAndCheckTopic = function (topic, pattern) {
        return checkTopic(normaliseTopic(topic), pattern);
    };

// Check if the supplied topic starts with the given pattern (both being normalised)
    const checkTopic = function (parts, pattern) {
        for (let i = 0; i < pattern.length; i++) {
            if (parts[i] !== pattern[i]) {
                return false;
            }
        }
        return true;
    };

// Show the connection status of the node depending on its underlying rfxtrx object
    const showConnectionStatus = function (node) {
        if (node.rfxtrx.connected === false) {
            node.status({fill: "red", shape: "ring", text: "disconnected"});
        } else {
            node.status({fill: "green", shape: "dot",
                text: "OK (v" + node.rfxtrx.firmwareVersion + " " + node.rfxtrx.firmwareType + ")"});
        }
    };

// Purge retransmission timers associated with this node and delete them to avoid leaking memory
    const purgeTimers = function (node) {
           let tx = {};
           for (tx in node.retransmissions) {
               if (node.retransmissions.hasOwnProperty(tx)) {
                   node.clearRetransmission(node.retransmissions[tx]);
                   delete node.retransmissions[tx];
               }
           }
       };

// The config node holding the PT2262 deviceList object
function RfxPT2262DeviceList(n) {
    RED.nodes.createNode(this, n);
    this.name = n.name;
    this.devices = n.devices;
}

// Register the PT2262 config node
RED.nodes.registerType("PT2262-device-list", RfxPT2262DeviceList);

// The config node holding raw output pulse time arrays
function RfxRawDeviceList(n) {
    RED.nodes.createNode(this, n);
    this.name = n.name;
    this.devices = n.devices;
}

// Register the raw times config node
RED.nodes.registerType("raw-device-list", RfxRawDeviceList);

// An input node for listening to messages from lighting remote controls
    function RfxLightsInNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        this.lighting1Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting1[evt.subtype] || "LIGHTING1_UNKNOWN") + "/" + evt.houseCode;
            if (evt.commandNumber === 5 || evt.commandNumber === 6) {
                msg.topic = msg.topic + "/0";
            } else {
                msg.topic = msg.topic + "/" + evt.unitCode;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                switch (evt.commandNumber) {
                    case 0 :
                    case 5 :
                        msg.payload = "Off";
                        break;

                    case 1 :
                    case 6 :
                        msg.payload = "On";
                        break;

                    case 2 :
                        msg.payload = "Dim";
                        break;

                    case 3 :
                        msg.payload = "Bright";
                        break;

                    case 7 :    // The ARC 'Chime' command - handled in rfx-doorbells so ignore it here
                        return;

                    default:
                        node.warn("rfx-lights-in: unrecognised Lighting1 command " + evt.commandNumber.toString(16));
                        return;
                }
                node.send(msg);
            }
        };
        this.lighting2Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting2[evt.subtype] || "LIGHTING2_UNKNOWN") + "/" + evt.id;
            if (evt.commandNumber > 2) {
                msg.topic = msg.topic + "/0";
            } else {
                msg.topic = msg.topic + "/" + evt.unitCode;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                switch (evt.commandNumber) {
                    case 0:
                    case 3:
                        msg.payload = "Off";
                        break;

                    case 1:
                    case 4:
                        msg.payload = "On";
                        break;

                    case 2:
                    case 5:
                        msg.payload = "Dim " + evt.level/15*100 + "%";
                        break;

                    default:
                        node.warn("rfx-lights-in: unrecognised Lighting2 command " + evt.commandNumber.toString(16));
                        return;
                }
                node.send(msg);
            }
        };
        this.lighting5Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting5[evt.subtype] || "LIGHTING5_UNKNOWN") + "/" + evt.id;
            if ((evt.commandNumber === 2 /* Group Off */ &&
                    (evt.subtype === 0 || evt.subtype === 2 || evt.subtype === 4 || evt.subtype === 17) ) ||
                (evt.commandNumber === 3) /* Group On */ &&
                    (evt.subtype === 2 || evt.subtype === 4 || evt.subtype === 17)) {
                msg.topic = msg.topic + "/0";
            } else {
                msg.topic = msg.topic + "/" + evt.unitCode;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                switch (evt.subtype) {
                    case 0: // Lightwave RF
                        switch (evt.commandNumber) {
                            case 0:
                            case 2:
                                msg.payload = "Off";
                                break;

                            case 1:
                                msg.payload = "On";
                                break;

                            case 3:
                            case 4:
                            case 5:
                            case 6:
                            case 7:
                                msg.payload = "Mood" + (evt.commandNumber - 2);
                                break;

                            case 16:
                                msg.payload = "Dim " + evt.level/31*100 + "%";
                                break;

                            case 17:
                            case 18:
                            case 19:
                                node.warn("Lighting5: LightwaveRF colour commands not implemented");
                                break;

                            default:
                                return;
                        }
                        break;

                    case 2:  // BBSB
                    case 4:  // CONRAD
                    case 17: // KANGTAI
                        switch (evt.commandNumber) {
                            case 0:
                            case 2:
                                msg.payload = "Off";
                                break;

                            case 1:
                            case 3:
                                msg.payload = "On";
                                break;

                            default:
                                return;
                        }
                        break;

                    case 6: // TRC02
                        switch (evt.commandNumber) {
                            case 0:
                                msg.payload = "Off";
                                break;

                            case 1:
                                msg.payload = "On";
                                break;

                            case 2:
                                msg.payload = "Bright";
                                break;

                            case 3:
                                msg.payload = "Dim";
                                break;

                            default:
                                node.warn("Lighting5: TRC02 colour commands not implemented");
                                return;
                        }
                        break;

                    case 13: // LEGRAND
                        switch (evt.commandNumber) {
                            case 0x00:
                                msg.payload = "Toggle";
                                break;

                            default:
                                return;
                        }
                        break;

                    default:
                        return;
                }
                node.send(msg);
            }
        };
        this.lighting6Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting6[evt.subtype] || "LIGHTING6_UNKNOWN") + "/" + evt.id + "/" + evt.groupCode;
            if (evt.commandNumber > 1) {
                msg.topic = msg.topic + "/0";
            } else {
                msg.topic = msg.topic + "/" + evt.unitCode;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                switch (evt.commandNumber) {
                    case 1:
                    case 3:
                        msg.payload = "Off";
                        break;

                    case 0:
                    case 2:
                        msg.payload = "On";
                        break;

                    default:
                        node.warn("rfx-lights-in: unrecognised Lighting6 command " + evt.commandNumber.toString(16));
                        return;
                }
                node.send(msg);
            }
        };
        this.security1Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi, battery: evt.batteryLevel}};
            if (evt.subtype === 2) {
                msg.topic = (rfxcom.security1[2]) + "/" + evt.id + "/";
                switch (evt.deviceStatus) {
                    case 0x10:
                        msg.payload = "Off";
                        msg.topic = msg.topic + "1";
                        break;

                    case 0x11:
                        msg.payload = "On";
                        msg.topic = msg.topic + "1";
                        break;

                    case 0x12:
                        msg.payload = "Off";
                        msg.topic = msg.topic + "2";
                        break;

                    case 0x13:
                        msg.payload = "On";
                        msg.topic = msg.topic + "2";
                        break;

                    default:
                        return;

                }
                if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                    node.send(msg);
                }
            }
        };
        this.hunterFanHandler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            if (evt.subtype === 0) {
                msg.topic = (rfxcom.hunterFan[evt.subtype] || "HUNTER_FAN_UNKNOWN") + "/" + evt.id;
                switch (evt.commandNumber) {
                    case 0x02:
                        msg.payload = "Toggle";
                        break;

                    default:
                        return;

                }
                if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                    node.send(msg);
                }
            }
        };
        this.fanHandler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.fan[evt.subtype] || "FAN_UNKNOWN") + "/" + evt.id;
            switch (evt.command) {
                case "Light":
                    msg.payload = "Toggle";
                    break;

                case "Light On":
                    msg.payload = "On";
                    break;

                case "Light Off":
                    msg.payload = "Off";
                    break;

                default:
                    return;

            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                node.send(msg);
            }
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("lighting1", node.lighting1Handler);
                        node.rfxtrx.removeListener("lighting2", node.lighting2Handler);
                        node.rfxtrx.removeListener("lighting5", node.lighting5Handler);
                        node.rfxtrx.removeListener("lighting6", node.lighting6Handler);
                        node.rfxtrx.removeListener("security1", node.security1Handler);
                        node.rfxtrx.removeListener("hunterfan", node.hunterFanHandler);
                        node.rfxtrx.removeListener("fan", node.fanHandler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("lighting1", this.lighting1Handler);
                node.rfxtrx.on("lighting2", this.lighting2Handler);
                node.rfxtrx.on("lighting5", this.lighting5Handler);
                node.rfxtrx.on("lighting6", this.lighting6Handler);
                node.rfxtrx.on("security1", this.security1Handler);
                node.rfxtrx.on("hunterfan", this.hunterFanHandler);
                node.rfxtrx.on("fan", this.fanHandler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-lights-in", RfxLightsInNode);

// An input node for listening to messages from PT622 (lighting4) devices
    function RfxPT2262InNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "all";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.devices = RED.nodes.getNode(n.deviceList).devices || [];
        this.ignoreUnmatched = n.ignoreUnmatched || false;
        this.squelchLevel = parseInt(n.squelchLevel) || 0;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        this.lighting4Handler = function (evt) {
            if (evt.rssi > node.squelchLevel) {
                let msg = {status: {rssi: evt.rssi}};
                let db = node.devices.filter(function (entry) {return parseInt(entry.rawData) === parseInt(evt.data)});
                if (db.length === 0) {
                    if (node.ignoreUnmatched === false) {
                        msg.raw = {data: evt.data, pulseWidth: evt.pulseWidth};
                        node.send(msg);
                    }
                } else {
                    if (node.topicSource === "all" || checkTopic(db[0].device, node.topic)) {
                        msg.topic = db[0].device.join("/");
                        msg.payload = db[0].payload;
                        node.send(msg);
                    }
                }
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("lighting4", node.lighting4Handler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("lighting4", this.lighting4Handler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-PT2262-in", RfxPT2262InNode);

// An output node for sending messages to lighting4 devices, using the PT2262/72 chips
function RfxPT2262OutNode(n) {
    RED.nodes.createNode(this, n);
    this.port = n.port;
    this.topicSource = n.topicSource || "msg";
    this.topic = stringToParts(n.topic);
    this.retransmit = n.retransmit || "none";
    this.retransmitInterval = n.retransmitInterval || 20;
    this.devices = RED.nodes.getNode(n.deviceList).devices || [];
    this.name = n.name;
    this.rfxtrxPort = RED.nodes.getNode(this.port);

    const node = this;
    node.retransmissions = {};
    node.mustDelete = false;
    if (node.retransmit === "once") {
        // Uses a Timeout: period in seconds
        node.setRetransmission = setTimeout;
        node.retransmitTime = 1000*node.retransmitInterval;
        node.clearRetransmission = clearTimeout;
        node.mustDelete = true;
    } else if (node.retransmit === "repeat") {
        // Uses an Interval: period in miutes
        node.setRetransmission = setInterval;
        node.retransmitTime = 60*1000*node.retransmitInterval;
        node.clearRetransmission = clearInterval;
    }

    if (node.rfxtrxPort) {
        node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
        if (node.rfxtrx !== null) {
            showConnectionStatus(node);
            getRfxcomSubtype(node.rfxtrx, "PT2262", ["lighting4"]);
            node.on("close", function () {
                purgeTimers(node);
                releasePort(node);
            });
            node.on("input", function (msg) {
                let db, topic, lastCommand, rawData = null, pulseWidth = null;
                // Get the topic from either the node or the message
                if (node.topicSource === "node" && node.topic !== undefined) {
                    topic = node.topic;
                } else if (msg.topic !== undefined) {
                    topic = stringToParts(msg.topic);
                }
                if (topic !== undefined && msg.payload !== undefined) {
                    // Lookup the topic/payload combination in the device list
                    db = node.devices.filter(function (entry) {
                        return msg.payload == entry.payload &&
                               topic.length === entry.device.length &&
                               checkTopic(topic, entry.device);
                    });
                    if (db.length >= 1) {
                        // If multiple matches use the first
                        rawData = db[0].rawData;
                        pulseWidth = db[0].pulseWidth;
                    } else {
                        node.warn("rfx-PT2262-out: no raw data found for '" + topic.join("/") + ":" + msg.payload + "'");
                    }
                } else if (msg.raw !== undefined && msg.raw.data !== undefined) {
                    // No topic or no payload or neither: check for raw data in the message
                    rawData = msg.raw.data;
                    if (msg.raw.pulseWidth !== undefined) {
                        pulseWidth = msg.raw.pulseWidth;
                    }
                    if (topic === undefined) {
                        topic = ["__none__"];
                    }
                }
                if (rawData !== null) {
                    try {
                        // Send the command for the first time
                        node.rfxtrx.transmitters['PT2262'].sendData(rawData, pulseWidth);
                        // If we reach this point, the command did not throw an error. Check if should retransmit
                        // it & set the Timeout or Interval as appropriate
                        if (node.retransmit !== "none") {
                            // Wrap the parseCommand arguments, and the retransmission key (=topic) in a
                            // function context, where the function lastCommand() can find them
                            topic = topic.join("/");
                            lastCommand = (function () {
                                let _rawData = rawData, _pulseWidth = pulseWidth, key = topic;
                                return function () {
                                    node.rfxtrx.transmitters['PT2262'].sendData(_rawData, _pulseWidth);
                                    if (node.mustDelete) {
                                        delete(node.retransmissions[key]);
                                    }
                                };
                            }());
                            if (node.retransmissions.hasOwnProperty(topic)) {
                                node.clearRetransmission(node.retransmissions[topic]);
                            }
                            node.retransmissions[topic] = node.setRetransmission(lastCommand, node.retransmitTime);

                        }
                    } catch (exception) {
                        node.warn(exception.message);
                    }
                }
            });
        }
    } else {
        node.error("missing config: rfxtrx-port");
    }
}

RED.nodes.registerType("rfx-PT2262-out", RfxPT2262OutNode);

// An output node for sending 'Raw' format messages
function RfxRawOutNode(n) {
    RED.nodes.createNode(this, n);
    this.port = n.port;
    this.topicSource = n.topicSource || "msg";
    this.topic = stringToParts(n.topic);
    this.retransmit = n.retransmit || "none";
    this.retransmitInterval = n.retransmitInterval || 20;
    this.devices = RED.nodes.getNode(n.deviceList).devices || [];
    this.name = n.name;
    this.rfxtrxPort = RED.nodes.getNode(this.port);

    const node = this;
    node.retransmissions = {};
    node.mustDelete = false;
    if (node.retransmit === "once") {
        // Uses a Timeout: period in seconds
        node.setRetransmission = setTimeout;
        node.retransmitTime = 1000*node.retransmitInterval;
        node.clearRetransmission = clearTimeout;
        node.mustDelete = true;
    } else if (node.retransmit === "repeat") {
        // Uses an Interval: period in miutes
        node.setRetransmission = setInterval;
        node.retransmitTime = 60*1000*node.retransmitInterval;
        node.clearRetransmission = clearInterval;
    }

    if (node.rfxtrxPort) {
        node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
        if (node.rfxtrx !== null) {
            showConnectionStatus(node);
            getRfxcomSubtype(node.rfxtrx, "RAW", ["rawtx"]);
            node.on("close", function () {
                purgeTimers(node);
                releasePort(node);
            });
            node.on("input", function (msg) {
                let db, topic, lastCommand, pulseTimes = null, repeats = null;
                // Get the topic from either the node or the message
                if (node.topicSource === "node" && node.topic !== undefined) {
                    topic = node.topic;
                } else if (msg.topic !== undefined) {
                    topic = stringToParts(msg.topic);
                }
                if (topic !== undefined && msg.payload !== undefined) {
                    // Lookup the topic/payload combination in the device list
                    db = node.devices.filter(function (entry) {
                        return msg.payload == entry.payload &&
                               topic.length === entry.device.length &&
                               checkTopic(topic, entry.device);
                    });
                    if (db.length >= 1) {
                        // If multiple matches use the first
                        pulseTimes = db[0].pulseTimes;
                        repeats = db[0].repeats;
                    } else {
                        node.warn("rfx-raw-out: no pulse data found for '" + topic.join("/") + ":" + msg.payload + "'");
                    }
                } else if (msg.raw !== undefined && msg.raw.data !== undefined) {
                    // No topic or no payload or neither: check for raw data in the message
                    pulseTimes = msg.raw.pulseTimes;
                    if (msg.raw.repeats !== undefined) {
                        repeats = msg.raw.repeats;
                    } else {
                        repeats = 5;
                    }
                    if (topic === undefined) {
                        topic = ["__none__"];
                    }
                }
                if (pulseTimes !== null) {
                     try {
                        // Send the command for the first time
                        let params = {pulseTimes: pulseTimes, repeats: repeats};
                        node.rfxtrx.transmitters['RAW'].sendMessage(null, params);
                        // If we reach this point, the command did not throw an error. Check if should retransmit
                        // it & set the Timeout or Interval as appropriate
                        if (node.retransmit !== "none") {
                            // Wrap the parseCommand arguments, and the retransmission key (=topic) in a
                            // function context, where the function lastCommand() can find them
                            topic = topic.join("/");
                            lastCommand = (function () {
                                // let _rawData = rawData, _pulseWidth = repeats, key = topic;
                                let _params = params, key = topic;
                                return function () {
                                    node.rfxtrx.transmitters['RAW'].sendMessage(null, _params);
                                    if (node.mustDelete) {
                                        delete(node.retransmissions[key]);
                                    }
                                };
                            }());
                            if (node.retransmissions.hasOwnProperty(topic)) {
                                node.clearRetransmission(node.retransmissions[topic]);
                            }
                            node.retransmissions[topic] = node.setRetransmission(lastCommand, node.retransmitTime);

                        }
                    } catch (exception) {
                        node.warn(exception.message);
                    }
                 }
            });
        }
    } else {
        node.error("missing config: rfxtrx-port");
    }
}

RED.nodes.registerType("rfx-raw-out", RfxRawOutNode);

// An input node for listening to messages from (mainly weather) sensors
    function RfxWeatherSensorNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        const sendWeatherMessage = function (evt, msg) {
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi};
                if (evt.hasOwnProperty("batteryLevel")) {
                    msg.status.battery = evt.batteryLevel;
                }
                msg.payload = {};
                if (evt.hasOwnProperty("temperature")) {
                    msg.payload.temperature = {value: evt.temperature, unit: "degC"};
                }
                if (evt.hasOwnProperty("setpoint")) {
                    msg.payload.setpoint = {value: evt.setpoint, unit: "degC"};
                    if (evt.hasOwnProperty("status")) {
                        msg.payload.status = evt.status;
                    }
                    if (evt.hasOwnProperty("mode")) {
                        msg.payload.mode = evt.mode;
                    }
                }
                if (evt.hasOwnProperty("barometer")) {
                    msg.payload.pressure = {value: evt.barometer, unit: "hPa"};
                }
                if (evt.hasOwnProperty("direction")) {
                    msg.payload.wind = {direction: {value: evt.direction, unit: "degrees"}};
                    if (evt.hasOwnProperty("averageSpeed")) {
                        msg.payload.wind.speed = {value: evt.averageSpeed, unit: "m/s"};
                        msg.payload.wind.gust = {value: evt.gustSpeed, unit: "m/s"};
                    } else {
                        msg.payload.wind.speed = {value: evt.gustSpeed, unit: "m/s"};
                    }
                    if (evt.hasOwnProperty("chillfactor")) {
                        msg.payload.wind.chillfactor = {value: evt.chillfactor, unit: "degC"};
                    }
                }
                if (evt.hasOwnProperty("humidity")) {
                    msg.payload.humidity = {
                        value:  evt.humidity,
                        unit:   "%",
                        status: rfxcom.humidity[evt.humidityStatus]
                    };
                }
                if (evt.hasOwnProperty("rainfall")) {
                    msg.payload.rainfall = {total: {value: evt.rainfall, unit: "mm"}};
                    if (evt.hasOwnProperty("rainfallRate")) {
                        msg.payload.rainfall.rate = {value: evt.rainfallRate, unit: "mm/hr"};
                    }
                } else if (evt.hasOwnProperty("rainfallIncrement")) {
                    msg.payload.rainfall = {increment: {value: evt.rainfallIncrement, unit: "mm"}};
                }
                if (evt.hasOwnProperty("uv")) {
                    msg.payload.uv = {value: evt.uv, unit: "UVIndex"};
                }
                if (evt.hasOwnProperty("forecast")) {
                    msg.payload.forecast = rfxcom.forecast[evt.forecast];
                }
                node.send(msg);
            }
        };
        this.bbq1Handler = function (evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.bbq1[evt.subtype] || "BBQ1_UNKNOWN") + "/" + evt.id})
        };
        this.temperaturerainHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.temperatureRain1[evt.subtype] || "TEMPERATURERAIN1_UNKNOWN") + "/" + evt.id})
        };
        this.temperatureHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.temperature1[evt.subtype] || "TEMPERATURE1_UNKNOWN") + "/" + evt.id})
        };
        this.humidityHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.humidity1[evt.subtype] || "HUMIDITY1_UNKNOWN") + "/" + evt.id})
        };
        this.temperaturehumidityHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.temperatureHumidity1[evt.subtype] || "TEMPERATUREHUMIDITY1_UNKNOWN") + "/" + evt.id})
        };
        this.temphumbaroHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.tempHumBaro1[evt.subtype] || "TEMPHUMBARO1_UNKNOWN") + "/" + evt.id})
        };
        this.thermostat1Handler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.thermostat1[evt.subtype] || "THERMOSTAT1_UNKNOWN") + "/" + evt.id})
        };
        this.rainHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.rain1[evt.subtype] || "RAIN1_UNKNOWN") + "/" + evt.id})
        };
        this.windHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.wind1[evt.subtype] || "WIND1_UNKNOWN") + "/" + evt.id})
        };
        this.uvHandler = function(evt) {
            sendWeatherMessage(evt, {topic: (rfxcom.uv1[evt.subtype] || "UV1_UNKNOWN") + "/" + evt.id})
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("bbq1", node.bbq1Handler);
                        node.rfxtrx.removeListener("temperaturerain1", node.temperaturerainHandler);
                        node.rfxtrx.removeListener("temperature1", node.temperatureHandler);
                        node.rfxtrx.removeListener("humidity1", node.humidityHandler);
                        node.rfxtrx.removeListener("temperaturehumidity1", node.temperaturehumidityHandler);
                        node.rfxtrx.removeListener("temphumbaro1", node.temphumbaroHandler);
                        node.rfxtrx.removeListener("thermostat1", node.thermostat1Handler);
                        node.rfxtrx.removeListener("rain1", node.rainHandler);
                        node.rfxtrx.removeListener("wind1", node.windHandler);
                        node.rfxtrx.removeListener("uv1", node.uvHandler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("bbq1", this.bbq1Handler);
                node.rfxtrx.on("temperaturerain1", this.temperaturerainHandler);
                node.rfxtrx.on("temperature1", this.temperatureHandler);
                node.rfxtrx.on("humidity1", this.humidityHandler);
                node.rfxtrx.on("temperaturehumidity1", this.temperaturehumidityHandler);
                node.rfxtrx.on("temphumbaro1", this.temphumbaroHandler);
                node.rfxtrx.on("thermostat1", this.thermostat1Handler);
                node.rfxtrx.on("rain1", this.rainHandler);
                node.rfxtrx.on("wind1", this.windHandler);
                node.rfxtrx.on("uv1", this.uvHandler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-sensor", RfxWeatherSensorNode);

// An input node for listening to messages from (electrical) energy & current monitors
    function RfxEnergyMeterNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        this.cartelectronicHandler = function (evt) {
            let msg = {topic: (rfxcom.cartelectronic[evt.subtype] || "CARTELECTRONIC_UNKNOWN") + "/" + evt.id};
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi, battery: evt.batteryLevel};
                msg.payload = {cartelectronic: {}};
                switch (evt.subtype) {
                    case 1 :
                        msg.payload.cartelectronic.type = "TIC";
                        msg.payload.cartelectronic.identifiantCompteur = evt.identifiantCompteur;
                        msg.payload.cartelectronic.typeContrat = evt.typeContrat;
                        msg.payload.cartelectronic.périodeTarifaireEnCours = evt.périodeTarifaireEnCours;
                        msg.payload.cartelectronic.compteur = evt.compteur;
                        msg.payload.cartelectronic.puissanceApparenteValide = evt.puissanceApparenteValide;
                        msg.payload.cartelectronic.puissanceApparente = evt.puissanceApparente;
                        msg.payload.cartelectronic.teleInfoPrésente = evt.teleInfoPrésente;
                        if (evt.hasOwnProperty("avertissemntJourEJP")) {
                            msg.payload.cartelectronic.avertissemntJourEJP = evt.avertissemntJourEJP;
                        }
                        if (evt.hasOwnProperty("avertissementCouleurDemain")) {
                            msg.payload.cartelectronic.avertissementCouleurDemain = evt.avertissementCouleurDemain;
                        }
                        msg.payload.counter = [{value: evt.compteur[0].valeur, unit: "Wh"}];
                        if (evt.compteur[1].période !== "non utilisé") {
                            msg.payload.counter.push({value: evt.compteur[1].valeur, unit: "Wh"});
                        }
                        break;

                    case 2 :
                        msg.payload.cartelectronic.type = "COMPTEUR";
                        msg.payload.cartelectronic.compteur = evt.compteur;
                        msg.payload.counter = [{value: evt.compteur[0].valeur, unit: "Count"},
                                               {value: evt.compteur[1].valeur, unit: "Count"}];
                        break;

                    case 3 :
                        msg.payload.cartelectronic.type = "LINKY";
                        msg.payload.cartelectronic.identifiantCompteur = evt.identifiantCompteur;
                        msg.payload.cartelectronic.compteur = evt.compteur;
                        msg.payload.cartelectronic.tensionMoyenne = evt.tensionMoyenne;
                        msg.payload.cartelectronic.puissanceApparenteValide = evt.puissanceApparenteValide;
                        msg.payload.cartelectronic.puissanceApparente = evt.puissanceApparente;
                        msg.payload.cartelectronic.teleInfoPrésente = evt.teleInfoPrésente;
                        msg.payload.cartelectronic.indexTariffaireEnCours = evt.indexTariffaireEnCours;
                        msg.payload.cartelectronic.avertissementCouleurAujourdHui = evt.avertissementCouleurAujourdHui;
                        msg.payload.cartelectronic.avertissementCouleurDemain = evt.avertissementCouleurDemain;

                        msg.payload.voltage = {value: evt.tensionMoyenne, unit: "V"};
                        msg.payload.counter = [{value: evt.compteur[0].valeur, unit: "Wh"}];
                        if (evt.compteur[1].contenu !== "non utilisé") {
                            msg.payload.counter.push({value: evt.compteur[1].valeur, unit: "Wh"});
                        }
                        break;

                    default:
                        msg.payload.cartelectronic.type = "INCONNU";
                        break;
                }
                node.send(msg);

            }
        };

        const sendMeterMessage = function (evt, packetType, msg) {
            // Discard all rfxMeter messages that are not simple counts
            if (packetType === 0x71 && evt.subtype !== 0) {
                return;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi};
                if (evt.hasOwnProperty("batteryLevel")) {
                    msg.status.battery = evt.batteryLevel;
                }
                msg.payload = {};
                if (evt.hasOwnProperty("voltage")) {
                    msg.payload.voltage = {value: evt.voltage, unit: "V"}
                }
                if (evt.hasOwnProperty("current")) {
                    msg.payload.current = {value: evt.current, unit: "A"}
                }
                if (evt.hasOwnProperty("power")) {
                    msg.payload.power = {value: evt.power, unit: "W"}
                }
                if (evt.hasOwnProperty("energy")) {
                    msg.payload.energy = {value: evt.energy, unit: "Wh"}
                }
                if (evt.hasOwnProperty("powerFactor")) {
                    msg.payload.powerFactor = {value: evt.powerFactor, unit: ""}
                }
                if (evt.hasOwnProperty("frequency")) {
                    msg.payload.frequency = {value: evt.frequency, unit: "Hz"}
                }
                if (evt.hasOwnProperty("counter")) {
                    msg.payload.counter = {value: evt.counter, unit: "Count"}
                }
                node.send(msg);
            }
        };
        this.elec1Handler = function (evt, packetType) {
            sendMeterMessage(evt, packetType, {topic: (rfxcom.elec1[evt.subtype] || "ELEC1_UNKNOWN") + "/" + evt.id})
        };
        this.elec23Handler = function (evt, packetType) {
            sendMeterMessage(evt, packetType, {topic: (rfxcom.elec23[evt.subtype] || "ELEC23_UNKNOWN") + "/" + evt.id})
        };
        this.elec4Handler = function (evt, packetType) {
            sendMeterMessage(evt, packetType, {topic: (rfxcom.elec4[evt.subtype] || "ELEC4_UNKNOWN") + "/" + evt.id})
        };
        this.elec5Handler = function (evt, packetType) {
            sendMeterMessage(evt, packetType, {topic: (rfxcom.elec5[evt.subtype] || "ELEC5_UNKNOWN") + "/" + evt.id})
        };
        this.rfxmeterHandler = function (evt, packetType) {
            sendMeterMessage(evt, packetType, {topic: (rfxcom.rfxMeter[evt.subtype] || "RFXMETER_UNKNOWN") + "/" + evt.id})
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("elec1", node.elec1Handler);
                        node.rfxtrx.removeListener("elec23", node.elec1Handler);
                        node.rfxtrx.removeListener("elec4", node.elec1Handler);
                        node.rfxtrx.removeListener("elec5", node.elec1Handler);
                        node.rfxtrx.removeListener("rfxmeter", node.rfxmeterHandler);
                        node.rfxtrx.removeListener("cartelectronic", node.cartelectronicHandler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("elec1", this.elec1Handler);
                node.rfxtrx.on("elec23", this.elec23Handler);
                node.rfxtrx.on("elec4", this.elec1Handler);
                node.rfxtrx.on("elec5", this.elec1Handler);
                node.rfxtrx.on("rfxmeter", this.rfxmeterHandler);
                node.rfxtrx.on("cartelectronic", this.cartelectronicHandler)
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-meter", RfxEnergyMeterNode);

// An input node for listening to messages from security and smoke detectors
    function RfxDetectorsNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.outputHeartbeats = n.outputHeartbeats || false;
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        node.heartbeats = {};
        node.HEARTBEATDELAY = []; // delay in minutes before declaring a detector has gone silent
        node.HEARTBEATDELAY[rfxcom.security1.X10_DOOR] = 90;
        node.HEARTBEATDELAY[rfxcom.security1.X10_PIR] = 90;
        node.HEARTBEATDELAY[rfxcom.security1.POWERCODE_DOOR] = 20;
        node.HEARTBEATDELAY[rfxcom.security1.POWERCODE_PIR] = 20;
        this.security1Handler = function (evt) {
            let msg = {topic: (rfxcom.security1[evt.subtype] || "SECURITY1_UNKNOWN") + "/" + evt.id};
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi};
                switch (evt.subtype) {
                    case rfxcom.security1.KD101:
                    case rfxcom.security1.SA30:
                    case rfxcom.security1.RM174RF:
                        if (evt.deviceStatus === rfxcom.security.PANIC) {
                            msg.payload = "Smoke";
                        }
                        break;

                    case rfxcom.security1.MEIANTECH:
                        if (evt.deviceStatus === rfxcom.security.PANIC) {
                            msg.status.battery = evt.batteryLevel;
                            msg.payload = "Motion";
                        }
                        break;

                    case rfxcom.security1.POWERCODE_AUX:
                        msg.status.battery = evt.batteryLevel;
                        msg.status.tampered = Boolean(evt.tampered);
                        if (evt.deviceStatus === rfxcom.security.ALARM) {
                            msg.payload = "Alarm";
                        } else if (evt.tampered) {
                            msg.payload = "Tamper";
                        }
                        break;

                    case rfxcom.security1.X10_SECURITY:
                        msg.status.battery = evt.batteryLevel;
                        switch (evt.deviceStatus) {
                            case 0x06:
                                msg.payload = "Panic";
                                break;

                            case 0x07:
                                msg.payload = "Cancel Panic";
                                break;

                            case 0x09:
                                msg.payload = "Arm Away";
                                break;

                            case 0x0a:
                                msg.payload = "Arm Away Delayed";
                                break;

                            case 0x0b:
                                msg.payload = "Arm Home";
                                break;

                            case 0x0c:
                                msg.payload = "Arm Home Delayed";
                                break;

                            case 0x0d:
                                msg.payload = "Disarm";
                                break;

                            default:
                                break;
                        }
                        break;

                    // These detectors are supposed to send "heartbeat" messages at more or less regular intervals
                    // However, some Chinese clones (e.g. Kerui P831) do not
                    case rfxcom.security1.POWERCODE_DOOR:
                    case rfxcom.security1.POWERCODE_PIR:
                    case rfxcom.security1.X10_DOOR:
                    case rfxcom.security1.X10_PIR:
                        msg.status.battery = evt.batteryLevel;
                        msg.status.tampered = Boolean(evt.tampered);
                        msg.status.state = evt.deviceStatus;
                        msg.status.delayed = Boolean(evt.deviceStatus === rfxcom.security.ALARM_DELAYED ||
                                                     evt.deviceStatus === rfxcom.security.NORMAL_DELAYED);
                        let deviceIsNotInAlarm = Boolean(evt.deviceStatus === rfxcom.security.NORMAL ||
                                                         evt.deviceStatus === rfxcom.security.NORMAL_DELAYED ||
                                                         evt.deviceStatus === rfxcom.security.NO_MOTION);
                        // Clear any existing heartbeat timeout & retrieve the device status from the last message
                        let lastDeviceStatus = NaN;
                        if (node.heartbeats.hasOwnProperty(msg.topic)) {
                            clearInterval(node.heartbeats[msg.topic].interval);
                            lastDeviceStatus = node.heartbeats[msg.topic].lastStatus;
                        }
                        // If this message has a NORMAL or NO_MOTION status, or there was an existing timer, set a new
                        // heartbeat timeout and record the current device status. Since Kerui P831 detectors and
                        // CreaJian-X21 switches (which use this packet type) send only MOTION (or ALARM) messages,
                        // they will never set a heartbeat timer
                        if (deviceIsNotInAlarm || isNaN(lastDeviceStatus) === false) {
                            node.heartbeats[msg.topic] = {
                                lastStatus: evt.deviceStatus,
                                interval:   (function () {
                                    const heartbeatStoppedMsg = {
                                        topic:                msg.topic,
                                        payload:              "Silent",
                                        lastMessageStatus:    msg.status,
                                        lastMessageTimestamp: Date.now(),
                                        lastHeardFrom:        new Date().toUTCString()
                                    };
                                    return setInterval(function () {
                                        delete heartbeatStoppedMsg._msgid;
                                        node.send(heartbeatStoppedMsg);
                                    }, 60*1000*node.HEARTBEATDELAY[evt.subtype]);
                                }())
                            };
                            // This ensures a clean shutdown on redeploy
                            node.heartbeats[msg.topic].interval.unref();
                            if (node.outputHeartbeats) {
                                msg.payload = "Heartbeat";
                            }
                        }
                        // Payload priority is Tamper > Alarm/Motion/Normal > Battery Low > Heartbeat (if enabled)
                        if (evt.batteryLevel === 0) {
                            msg.payload = "Battery Low";
                        }
                        if (evt.deviceStatus === rfxcom.security.MOTION) {
                            msg.payload = "Motion";
                        } else if (evt.deviceStatus !== lastDeviceStatus) {
                            // Only report ALARM/NORMAL status (i.e. window open/closed) if the status has changed
                            if (evt.deviceStatus === rfxcom.security.ALARM ||
                                evt.deviceStatus === rfxcom.security.ALARM_DELAYED) {
                                msg.payload = "Alarm";
                            } else if (evt.deviceStatus === rfxcom.security.NORMAL ||
                                       evt.deviceStatus === rfxcom.security.NORMAL_DELAYED) {
                                msg.payload = "Normal";
                            }
                        }
                        if (evt.tampered) {
                            msg.payload = "Tamper";
                        }
                        break;

                    default:
                        break;
                }
                if (msg.payload) {
                    node.send(msg);
                }
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    node.rfxtrx.removeListener("security1", node.security1Handler);
                    let heartbeat = {};
                    for (heartbeat in node.heartbeats) {
                        if (node.heartbeats.hasOwnProperty(heartbeat)) {
                            clearInterval(node.heartbeats[heartbeat]);
                        }
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("security1", this.security1Handler)
            }
        }
    }

    RED.nodes.registerType("rfx-detector-in", RfxDetectorsNode);

// An output node for sending PANIC messages to K101 type smoke detectors (sound the alarm!)
    function RfxAlarmOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;


        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // any message, regardless of payload, sends a PANIC status to the address
                    let path = [], protocolName, subtype = -1, deviceAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-alarm-out ") + ": missing topic");
                        return;
                    }
                    deviceAddress = path.slice(1, 2);
                    try {
                        protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["security1"]);
                        if (subtype < 0 || ["X10_SECURITY", "KD101", "MEIANTECH", "SA30", "RM174RF"].indexOf(protocolName) < 0) {
                            node.warn((node.name || "rfx-alarm-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            let command = rfxcom.security.PANIC;
                            if (protocolName === "X10_SECURITY") {
                                if (/disarm/i.test(msg.payload)) {
                                    command = rfxcom.security.DISARM;
                                } else if (/don't|cancel|end|stop/i.test(msg.payload)) {
                                    command = rfxcom.security.END_PANIC;
                                } else if (/arm.*home.*delay/i.test(msg.payload)) {
                                    command = rfxcom.security.ARM_HOME_DELAYED;
                                } else if (/arm.*home/i.test(msg.payload)) {
                                    command = rfxcom.security.ARM_HOME;
                                } else if (/arm.*delay/i.test(msg.payload)) {
                                    command = rfxcom.security.ARM_AWAY_DELAYED;
                                } else if (/arm/i.test(msg.payload)) {
                                    command = rfxcom.security.ARM_AWAY;
                                } else if (/panic/i.test(msg.payload) === false) {
                                    return;
                                }
                            }
                            node.rfxtrx.transmitters[protocolName].sendStatus(deviceAddress, command);
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-alarm-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-alarm-out", RfxAlarmOutNode);

// An output node for sending messages to light switches & dimmers (including most types of plug-in switch)
    function RfxLightsOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.retransmit = n.retransmit || "none";
        this.retransmitInterval = n.retransmitInterval || 20;
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        node.retransmissions = {};
        node.mustDelete = false;
        if (node.retransmit === "once") {
            // Uses a Timeout: period in seconds
            node.setRetransmission = setTimeout;
            node.retransmitTime = 1000*node.retransmitInterval;
            node.clearRetransmission = clearTimeout;
            node.mustDelete = true;
        } else if (node.retransmit === "repeat") {
            // Uses an Interval: period in minutes
            node.setRetransmission = setInterval;
            node.retransmitTime = 60*1000*node.retransmitInterval;
            node.clearRetransmission = clearInterval;
        }

        // Parse the message payload to obtain the normalised representation of the 'level' associated with a dimming
        // command. The result is either an integer in the range levelRange[0]..levelRange[1], '+' (meaning increase
        // the brightness level), or '-' (meaning reduce it). An input numeric value should be in the range 0..1, or
        // equivalently in the range 0%..100%, unless it is associated with an alexaCommand. In that case, an input
        // value in the range 0 to 100 is interpreted as a percentage (SetPercentageRequest). The value supplied with
        // IncrementPercentageRequest & DecrementPercentageRequest is ignored.
        // An empty level range means this device accepts Dim/Bright or Dim-/Dim+ commands only
        // Called only from parseCommand
        const parseDimLevel = function (payload, alexaCommand, levelRange) {
            let value = NaN;
            if (typeof alexaCommand === "string") {
                if (alexaCommand === "SetPercentageRequest") {
                    value = (payload.value || payload)/100;
                    if (levelRange.length === 0) {
                        return value >= 0.5 ? "+" : "-"
                    }
                } else if (alexaCommand === "IncrementPercentageRequest") {
                    return "+"
                } else if (alexaCommand === "DecrementPercentageRequest") {
                    return "-"
                }
            } else {
                if (levelRange.length === 0 || /[0-9]+/.test(payload) === false) {
                    if (/dim.*\+/i.test(payload)) { // 'dim+' means 'bright'
                        return "+";
                    } else if (/dim/i.test(payload)) {
                        return "-";
                    } else if (/bright.*-/i.test(payload)) { // 'bright-' means 'dim'
                        return "-";
                    } else if (/bright/i.test(payload)) {
                        return "+";
                    }
                }
                if (/[0-9]+/.test(payload) === false) {
                    if (/\+/.test(payload)) {
                        return "+";
                    } else if (/-/.test(payload)) {
                        return "-";
                    }
                }
                const match = /[0-9]+(\.[0-9]*)?/.exec(payload);
                if (match !== null) {
                    value = parseFloat(match[0]);
                }
                if (payload.match(/[0-9] *%/)) {
                    value = value/100;
                }
            }
            value = Math.max(0, Math.min(1, value));
            if (levelRange === undefined) {
                return NaN;
            } else {
                return Math.round(levelRange[0] + value*(levelRange[1] - levelRange[0]));
            }
        };

        // Parse the message payload to obtain the room number, if present. Returns an object holding the room number if
        // found, (defaulting to 1 if not found,) and the contents of the payload with the interpreted part removed
        const parseRoomNumber = function (payload) {
            let roomNumber = 1;
            const match = /room *([0-9]+)/i.exec(payload);
            if (match !== null && match.length >= 2) {
                roomNumber = parseInt(match[1]);
                payload = payload.replace(match[0], "");
            }
            return {payload: payload, number: roomNumber};
        };

        // Parse the message payload looking for lighting command messages (including Alexa-home command messages), calling
        // the corresponding function in the node-rfxcom API to implement it. All parameter checking is delegated to
        // that API. If no valid command is recognised, this function warns the user, but if the transmitter throws an
        // Error this function does not catch it
        const parseCommand = function (protocolName, address, payload, alexaCommand, levelRange) {
            if (/dim|bright|level|%|[0-9]\.|\.[0-9]/i.test(payload) || /Percentage/i.test(alexaCommand)) {
                const room = parseRoomNumber(payload);
                const level = parseDimLevel(room.payload, alexaCommand, levelRange);
                if (isFinite(level)) {
                    node.rfxtrx.transmitters[protocolName].setLevel(address, level);
                } else if (level === '+') {
                    node.rfxtrx.transmitters[protocolName].increaseLevel(address, room.number);
                } else if (level === '-') {
                    node.rfxtrx.transmitters[protocolName].decreaseLevel(address, room.number);
                } else {
                    node.warn("Don't understand dimming command '" + payload + "'");
                }
            } else if (/on/i.test(payload) || payload === 1 || payload === true) {
                if (node.rfxtrx.transmitters[protocolName].isSubtype("X10_SECURITY")) {
                    node.rfxtrx.transmitters[protocolName].switchLightOn(address[0], address[1]);
                } else if (node.rfxtrx.transmitters[protocolName].isSubtype("FALMEC")) {
                    node.rfxtrx.transmitters[protocolName].switchLightOn(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].switchOn(address);
                }
            } else if (/off/i.test(payload) || payload === 0 || payload === false) {
                if (node.rfxtrx.transmitters[protocolName].isSubtype("X10_SECURITY")) {
                    node.rfxtrx.transmitters[protocolName].switchLightOff(address[0], address[1]);
                } else if (node.rfxtrx.transmitters[protocolName].isSubtype("FALMEC")) {
                    node.rfxtrx.transmitters[protocolName].switchLightOff(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].switchOff(address);
                }
            } else if (/mood/i.test(payload)) {
                const match = /mood *([0-9]+)/i.exec(payload);
                if (match !== null && match.length >= 2) {
                    const mood = parseInt(match[1]);
                    node.rfxtrx.transmitters[protocolName].setMood(address, mood);
                } else {
                    node.warn("Missing mood number");
                }
            } else if (/scene/i.test(payload)) {
                const room = parseRoomNumber(payload);
                const match = /scene *([0-9]+)/i.exec(room.payload);
                if (match !== null && match.length >= 2) {
                    const sceneNumber = parseInt(match[1]);
                    node.rfxtrx.transmitters[protocolName].setScene(address, sceneNumber, room.number);
                } else {
                    node.warn("Missing scene number");
                }
            } else if (/toggle/i.test(payload)) {
                if (node.rfxtrx.transmitters[protocolName].isSubtype(["HUNTER_FAN", "SIEMENS_SF01", "LUCCI_AIR",
                        "WESTINGHOUSE_7226640", "CASAFAN", "LUCCI_AIR_DC", "FT1211R", "LUCCI_AIR_DCII", "NOVY"])) {
                    node.rfxtrx.transmitters[protocolName].toggleLightOnOff(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].toggleOnOff(address);
                }
            } else if (/program|learn|pair/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].program(address);
            } else {
                node.warn("Don't understand payload '" + payload + "'");
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    purgeTimers(node);
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, unitAddress, lastCommand, topic;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn("rfx-lights-out: missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["lighting1", "lighting2", "lighting3",
                                                  "lighting5", "lighting6", "homeConfort", "security1", "fan", "hunterFan"]);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-lights-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            switch (node.rfxtrx.transmitters[protocolName].packetType) {
                                case "hunterFan":
                                case "fan":
                                    deviceAddress = path[1];
                                    unitAddress = [];
                                    break;

                                default:
                                    deviceAddress = path.slice(1, -1);
                                    unitAddress = parseUnitAddress(path.slice(-1)[0]);
                                    break;
                            }
                            let levelRange = [];
                            switch (node.rfxtrx.transmitters[protocolName].packetType) {
                                case "lighting1" :
                                case "lighting6" :
                                case "homeConfort" :
                                case "fan":
                                case "hunterFan":
                                    break;

                                case "lighting2" :
                                    if (!node.rfxtrx.transmitters[protocolName].isSubtype("KAMBROOK")) {
                                        levelRange = [0, 15];
                                    }
                                    break;

                                case "lighting3" :
                                    levelRange = [0, 10];
                                    break;

                                case "lighting5" :
                                    if (node.rfxtrx.transmitters[protocolName].isSubtype("LIGHTWAVERF")) {
                                        levelRange = [0, 31];
                                    } else if (node.rfxtrx.transmitters[protocolName].isSubtype("IT")) {
                                        levelRange = [1, 8];
                                    } else if (node.rfxtrx.transmitters[protocolName].isSubtype(["MDREMOTE", "MDREMOTE_108"])) {
                                        levelRange = [1, 3];
                                    } else if (node.rfxtrx.transmitters[protocolName].isSubtype("MDREMOTE_107")) {
                                        levelRange = [1, 6];
                                    }
                                    break;
                            }
                            try {
                                // Send the command for the first time
                                parseCommand(protocolName, deviceAddress.concat(unitAddress), msg.payload, msg.command, levelRange);
                                // If we reach this point, the command did not throw an error. Check if we should
                                // retransmit it & set the Timeout or Interval as appropriate
                                if (node.retransmit !== "none") {
                                    topic = path.join("/");
                                    // Wrap the parseCommand arguments, and the retransmission key (=topic) in a
                                    // function context, where the function lastCommand() can find them
                                    lastCommand = (function () {
                                        let protocol = protocolName, address = deviceAddress.concat(unitAddress),
                                            payload = msg.payload, alexaCommand = msg.command,
                                            range = levelRange, key = topic;
                                        return function () {
                                            parseCommand(protocol, address, payload, alexaCommand, range);
                                            if (node.mustDelete) {
                                                delete(node.retransmissions[key]);
                                            }
                                        };
                                    }());
                                    if (node.retransmissions.hasOwnProperty(topic)) {
                                        node.clearRetransmission(node.retransmissions[topic]);
                                    }
                                    node.retransmissions[topic] = node.setRetransmission(lastCommand, node.retransmitTime);
                                }
                            } catch (exception) {
                                if (exception.message.indexOf("is not a function") >= 0) {
                                    const alexaCommand = typeof msg.command === "string" ? msg.command + ":" : "";
                                    node.warn("Input '" +  alexaCommand + msg.payload + "' generated command '" +
                                        exception.message.match(/[^_a-zA-Z]([_0-9a-zA-Z]*) is not a function/)[1] + "' not supported by device");
                                } else {
                                    node.warn((node.name || "rfx-lights-out") + ": " + exception);
                                }
                            }
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-lights-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-lights-out", RfxLightsOutNode);

// An output node for sending messages to fans
    function RfxFanOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        const parseSpeed = function (payload, speedRange) {
            let value = NaN;
            if (speedRange.length === 0 || /[0-9]+/.test(payload) === false) {
                if (/decr|redu|slow/i.test(payload)) {
                    return "-";
                } else if (/incr|fast/i.test(payload)) {
                    return "+";
                }
            }
            if (/[0-9]+/.test(payload) === false) {
                if (/\+/.test(payload)) {
                    return "+";
                } else if (/-/.test(payload)) {
                    return "-";
                }
            }
            if (speedRange === [1, 3]){
                if (/lo/i.test(payload)) {
                    return 1;
                } else if(/med/i.test(payload)) {
                    return 2;
                } else if(/hi/i.test(payload)) {
                    return 3;
                }
            }
            if (speedRange === [1, 4]) {
                if (/ful/i.test(payload)) {
                    return 4;
                }
            }
            const numericStrings = /[0-9]+(\.[0-9]*)?/.exec(payload);
            if (numericStrings === null || speedRange === undefined) {
                return NaN;
            }
            value = parseFloat(numericStrings[0]);
            if (payload.match(/[0-9] *%/)) {
                value = value/100;
                value = Math.max(0, Math.min(1, value));
                if (speedRange[0] === 0) {
                    return Math.round(value*speedRange[1]);
                } else {
                    return Math.max(1, Math.round(value*speedRange[1]));
                }
            } else {
                value = Math.round(value);
                if (value > speedRange[1]) {
                    return speedRange[1];
                } else if (value < speedRange[0]) {
                    return speedRange[0];
                } else {
                    return value;
                }
            }
        };

        const parseTime = function (payload, protocolName) {
            const numericStrings = /[0-9]+(\.[0-9]*)?/.exec(payload);
            if (numericStrings === null) {
                return 1;
            }
            const value = parseFloat(numericStrings[0]);
            return Math.round(value);
        };

        const parseCommand = function (protocolName, address, payload, speedRange) {
            if (/^ *\+ *$|^ *- *$|speed|slo|fast|lo|med|hi|ful|%|[0-9]\.|\.[0-9]/i.test(payload)) {
                const speed = parseSpeed(payload, speedRange);
                if (isFinite(speed)) {
                    node.rfxtrx.transmitters[protocolName].setSpeed(address, speed);
                } else if (speed === '+') {
                    node.rfxtrx.transmitters[protocolName].increaseSpeed(address);
                } else if (speed === '-') {
                    node.rfxtrx.transmitters[protocolName].decreaseSpeed(address);
                } else {
                    node.warn("Don't understand speed-set command '" + payload + "'");
                }
            } else if (/power|toggle/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].toggleOnOff(address);
            } else if (/time/i.test(payload)) {
                const timeValue = parseTime(payload, protocolName);
                if (protocolName === "NOVY") {
                    node.rfxtrx.transmitters[protocolName].toggleOnOff(address);
                    node.rfxtrx.transmitters[protocolName].toggleOnOff(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].startTimer(address, timeValue);
                }
            } else if (/on|start|run/i.test(payload) || payload === 1 || payload === "1" || payload === true) {
                try {
                    node.rfxtrx.transmitters[protocolName].toggleOnOff(address);
                } catch {
                    if (speedRange.length > 0) {
                        node.rfxtrx.transmitters[protocolName].setSpeed(address, 1);
                    } else {
                        node.rfxtrx.transmitters[protocolName].increaseSpeed(address);
                    }
                }
            } else if (/off|stop/i.test(payload) || payload === 0 || payload === "0" || payload === false) {
                node.rfxtrx.transmitters[protocolName].switchOff(address);
            } else if (/program|learn|pair/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].program(address);
            } else if (/reverse/i.test(payload)) {
                try {
                    node.rfxtrx.transmitters[protocolName].setFanDirection(address, 0);
                } catch {
                    node.rfxtrx.transmitters[protocolName].toggleFanDirection(address);
                }
            } else if (/natural|normal|forward/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].setFanDirection(address, 1);
            } else {
                node.warn("Don't understand payload '" + payload + "'");
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, speedRange;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-fan-out ") + ": missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    deviceAddress = path.slice(1, 2);
                    deviceAddress = path[1];
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["fan", "hunterFan"]);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-fan-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            switch (protocolName) {
                                case "ITHO_CVE_RFT":
                                    speedRange = [1, 3];
                                    break;

                                case "LUCCI_AIR":
                                case "WESTINGHOUSE_7226640":
                                case "CASAFAN":
                                case "HUNTER_FAN":
                                    speedRange = [0, 3];
                                    break;

                                case "FT1211R":
                                    speedRange = [1, 5];
                                    break;

                                case "FALMEC":
                                    speedRange = [0, 4];
                                    break;

                                case "LUCCI_AIR_DCII":
                                    speedRange = [0, 6];
                                    break;

                                case "ITHO_CVE_ECO_RFT":
                                    speedRange = [1, 4];
                                    break;

                                default:
                                    speedRange = [];
                                    break;
                            }
                            parseCommand(protocolName, deviceAddress, msg.payload, speedRange);
                        }
                    } catch (exception) {
                        if (exception.message.indexOf("is not a function") >= 0) {
                            const alexaCommand = typeof msg.command === "string" ? msg.command + ":" : "";
                            node.warn("Input '" +  alexaCommand + msg.payload + "' generated command '" +
                                exception.message.match(/[^_a-zA-Z]([_0-9a-zA-Z]*) is not a function/)[1] + "' not supported by device");
                        } else {
                            node.warn((node.name || "rfx-fan-out") + ": " + exception);
                        }
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-fan-out", RfxFanOutNode);

// An input node for listening to messages from doorbells
    function RfxDoorbellInNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        this.lighting1Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting1[evt.subtype] || "LIGHTING1_UNKNOWN") + "/" + evt.houseCode + "/" + evt.unitCode;
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                if (evt.subtype !== 0x01 || evt.commandNumber !== 7) {
                    return;
                }
                node.send(msg);
            }
        };
        this.chime1Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.chime1[evt.subtype] || "CHIME1_UNKNOWN") + "/" + evt.id;
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                if (evt.subtype === rfxcom.chime1.BYRON_SX) {
                    msg.payload = evt.commandNumber;
                }
                node.send(msg);
            }
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("lighting1", node.lighting1Handler);
                        node.rfxtrx.removeListener("chime1", node.chime1Handler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("lighting1", this.lighting1Handler);
                node.rfxtrx.on("chime1", this.chime1Handler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-doorbell-in", RfxDoorbellInNode);

// An output node for sending messages to doorbells
    function RfxDoorbellOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        // Generate the chime command depending on the subtype and tone parameter, if any
        const parseCommand = function (protocolName, address, str) {
            let sound = NaN;
            if (str !== undefined) {
                sound = parseInt(str);
            }
            try {
                if (protocolName === "BYRON_SX" && !isNaN(sound)) {
                    node.rfxtrx.transmitters[protocolName].chime(address, sound);
                } else {
                    node.rfxtrx.transmitters[protocolName].chime(address);
                }
            } catch (exception) {
                node.warn(exception);
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, unitAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-doorbell-out ") + ": missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    deviceAddress = path.slice(1, 2);
                    if (protocolName === 'ARC') {
                        unitAddress = parseUnitAddress(path.slice(-1)[0]);
                    } else {
                        unitAddress = [];
                    }
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["lighting1", "chime1"]);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-doorbell-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            parseCommand(protocolName, deviceAddress.concat(unitAddress), msg.payload);
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-doorbell-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-doorbell-out", RfxDoorbellOutNode);

// An output node for sending messages to Smartwares TRVs
    function RfxTRVOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        // Generate the device command depending on the message payload and/or command
        const parseCommand = function (protocolName, address, msg) {
            let temperature = NaN;
            // Check for an Alexa command
            if (msg.hasOwnProperty("command")) {
                if (msg.command === "SetTargetTemperatureRequest") {
                    temperature = msg.payload;
                } else if (msg.command === "TurnOnRequest") {
                    msg.payload = "ON'";
                } else if (msg.command === "TurnOffRequest") {
                    msg.payload = "OFF";
                } else {
                    node.warn("Unsupported Alexa command: " + msg.command)
                }
            } else if (typeof msg.payload === "string") {
                temperature = parseFloat(msg.payload);
            } else if (typeof msg.payload === "number") {
                temperature = msg.payload;
            }
            try {
                if (!isNaN(temperature)){
                    node.rfxtrx.transmitters[protocolName].setTemperature(address, temperature);
                } else if (/day|on|normal|heat/i.test(msg.payload)) {
                    node.rfxtrx.transmitters[protocolName].setDayMode(address);
                } else if (/night|off|setback|away/i.test(msg.payload)) {
                    node.rfxtrx.transmitters[protocolName].setNightMode(address);
                }
            } catch (exception) {
                node.warn(exception);
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, unitAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-trv-out ") + ": missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    deviceAddress = path.slice(1, 2);
                    if (protocolName === 'SMARTWARES') {
                        unitAddress = parseUnitAddress(path.slice(-1)[0]);
                    } else {
                        unitAddress = [];
                    }
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["radiator1"]);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-trv-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            parseCommand(protocolName, deviceAddress.concat(unitAddress), msg);
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-trv-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-trv-out", RfxTRVOutNode);

// An output node for sending messages to stoves and heating controllers
    function RfxHeaterOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        // Call the device command function depending on the message payload and/or command
        const parseCommand = function (protocolName, address, msg) {
            let percentage = NaN, command = msg.payload;
            // Check for an Alexa command
            if (msg.hasOwnProperty("command")) {
                if (msg.command === "SetPercentageRequest") {
                    percentage = msg.payload;
                    command = "SET";
                } else if (msg.command === "IncrementPercentageRequest") {
                    percentage = msg.payload;
                    command = "UP";
                } else if (msg.command === "DecrementPercentageRequest") {
                    percentage = msg.payload;
                    command = "DOWN";
                } else if (msg.command === "TurnOnRequest") {
                    command = "ON";
                } else if (msg.command === "TurnOffRequest") {
                    command = "OFF";
                } else {
                    command = "Alexa: " + msg.command;
                }
            } else if (typeof msg.payload === "number") {
                // A number 1-5 is translated to a percentage 20% - 100%, 0 means OFF
                // This is to suit the MCZ devices 'flame power' range
                if (msg.payload <= 0) {
                    command = "OFF";
                } else if (!isNaN(msg.payload)) {
                    percentage = msg.payload*20;
                    command = "SET";
                }
            } else if (typeof msg.payload === "boolean") {
                // Boolean means ON or OFF
                command = msg.payload ? "ON" : "OFF";
            } else if (typeof msg.payload === "string") {
                // If the string contains a number, it is either a burner number (1-2), a level (1-5), or a percentage (20%-100%)
                const numericField = /[+-]?(([0-9]+\.[0-9]*)|(\.[0-9]+)|([0-9]+))(?![0-9])/.exec(msg.payload);
                if (numericField !== null) {
                    percentage = parseFloat(numericField[0]);
                    // Check for the Mertik G6R-H4TB burner number
                    if (/ON|HEA/i.test(msg.payload)) {
                        switch (percentage) {
                            case 1:
                                command = "ON";
                                break;
                            case 2:
                                command = "ON2";
                                break;
                            default:
                                command = msg.payload;
                        }
                    } else if (/OFF/i.test(msg.payload)) {
                        switch (percentage) {
                            case 1:
                                command = "OFF";
                                break;
                            case 2:
                                command = "OFF2";
                                break;
                            default:
                                command = msg.payload;
                        }
                    } else {
                        if (!msg.payload.match(/[0-9.] *%/)) {
                        percentage = percentage*20;
                        }
                        // Zero % means 'off'
                        if (percentage === 0) {
                            command = "OFF";
                        } else {
                            command = "SET";
                        }
                    }
                // Normalise remaining string commands
                } else if (/ON|HEA/i.test(msg.payload)) {
                    command = "ON";
                } else if (/COO/i.test(msg.payload)) {
                    command = "COOL";
                } else if (/OFF/i.test(msg.payload)) {
                    command = "OFF";
                } else if (/RUN *UP/i.test(msg.payload)) {
                    command = "RUNUP";
                } else if (/RUN *DOW/i.test(msg.payload)) {
                    command = "RUNDOWN";
                } else if (/STO/i.test(msg.payload)) {
                    command = "STOP";
                } else if (/UP|INC|\+/i.test(msg.payload)) {
                    command = "UP";
                } else if (/DOW|DEC|-/i.test(msg.payload)) {
                    command = "DOWN";
                } else if (/program|learn|pair/i.test(msg.payload)) {
                    command = "PROGRAM";
                }
            }

            // If msg.payload is an object, it must be interpreted depending on the packet type
            if (/MCZ/i.test(protocolName)) {
                // We expect some or all of:
                //    fanSpeed - a single number applies to all fans, or an array
                //    mode (or command) - "On", "Man", "Auto", or "Eco"
                //    beep - true or false
                //    flamePower - 1 to 5
                let params = {};
                if (msg.payload.hasOwnProperty("beep")) {
                    params.beep = msg.payload.beep;
                }
                if (msg.payload.hasOwnProperty("fanSpeed")) {
                    if (typeof msg.payload.fanSpeed === "number") {
                        params.fanSpeed = [msg.payload.fanSpeed, msg.payload.fanSpeed, msg.payload.fanSpeed];
                    } else {
                        params.fanSpeed = msg.payload.fanSpeed;
                    }
                }
                if (msg.payload.hasOwnProperty("flamePower")) {
                    params.flamePower = msg.payload.flamePower;
                }
                if (msg.payload.hasOwnProperty("mode")) {
                    params.mode = msg.payload.mode;
                } else if (msg.payload.hasOwnProperty("command")) {
                    params.mode = msg.payload.command;
                }
                if ("flamePower" in params) {
                    if ("mode" in params === false) {
                        params.mode = "Man";
                    }
                // Handle the commands generated by a msg.payload which is not an object
                } else if ("mode" in params) {
                    params.flamePower = 1;
                } else if (command === "OFF") {
                    params.mode = "Off";
                    params.flamePower = 1;
                } else if (command === "ON") {
                    params.mode = "Auto";
                    params.flamePower = 1;
                } else if (command === "SET") {
                    // "set 0%" means turn off (e.g. from an Alexa command)
                    if (percentage <= 0) {
                        params.mode = "Off";
                        params.flamePower = 1;
                    } else {
                        params.mode = "Man";
                        params.flamePower = Math.max(percentage/20, 1);
                    }
                }
                if ("mode" in params === false || "flamePower" in params === false) {
                    if (typeof command === "string") {
                        node.warn("Unsupported command '" + command + "'");
                    } else {
                        node.warn("MCZ: missing mode/flamePower parameter in message object " + JSON.stringify(command));
                    }
                } else {
                    try {
                        node.rfxtrx.transmitters[protocolName].sendMessage(address, params);
                    } catch (exception) {
                        node.warn(exception);
                    }
                }
            } else if (/DIGIMAX/i.test(protocolName)) {
                let params = {setpoint: 20, mode: 0, temperature:20};
                // The only essential object parameter is status, the rest are passed on if they exist
                if (msg.payload.hasOwnProperty("status")) {
                    params.status = msg.payload.status;
                    if (msg.payload.hasOwnProperty("setpoint")) {
                        params.setpoint = msg.payload.setpoint.value;
                    }
                    if (msg.payload.hasOwnProperty("mode")) {
                        params.mode = msg.payload.mode;
                    }
                    if (msg.payload.hasOwnProperty("temperature")) {
                        params.temperature = msg.payload.temperature.value;
                    }
                } else if (command === "ON") {
                    params.status = 1;
                } else if (command === "OFF") {
                    params.status = 2;
                } else if (command === "COOL") {
                    params.mode = 1;
                    params.status = 1;
                } else if (command === "SET") {
                    if (percentage > 0) {
                        params.status = 1;
                    } else {
                        params.status = 2;
                    }
                }
                if (params.status === undefined) {
                    node.warn("DIGIMAX: missing status parameter in message object " + JSON.stringify(command));
                } else {
                    try {
                        node.rfxtrx.transmitters[protocolName].sendMessage(address, params);
                    } catch (exception) {
                        node.warn(exception);
                    }
                }
            } else {
                // Other device types don't accept object payloads, commands only
                try {
                    switch (command) {
                        case "ON":
                            node.rfxtrx.transmitters[protocolName].switchOn(address);
                            break;

                        case "OFF":
                            node.rfxtrx.transmitters[protocolName].switchOff(address);
                            break;

                        case "ON2":
                            node.rfxtrx.transmitters[protocolName].switchOn2(address);
                            break;

                        case "OFF2":
                            node.rfxtrx.transmitters[protocolName].switchOff2(address);
                            break;

                        case "UP":
                            node.rfxtrx.transmitters[protocolName].up(address);
                            break;

                        case "DOWN":
                            node.rfxtrx.transmitters[protocolName].down(address);
                            break;

                        case "RUNUP":
                            node.rfxtrx.transmitters[protocolName].runUp(address);
                            break;

                        case "RUNDOWN":
                            node.rfxtrx.transmitters[protocolName].runDown(address);
                            break;

                        case "STOP":
                            node.rfxtrx.transmitters[protocolName].stop(address);
                            break;

                        case "PROGRAM":
                            node.rfxtrx.transmitters[protocolName].program(address);
                            break;

                        case "SET":
                            if (percentage > 0) {
                                node.rfxtrx.transmitters[protocolName].switchOn(address);
                            } else {
                                node.rfxtrx.transmitters[protocolName].switchOff(address);
                            }
                            break;

                        default:
                            if (typeof command === "string") {
                                node.warn("Unsupported command '" + command + "'");
                            } else {
                                node.warn("Don't understand message object " + JSON.stringify(command));
                            }
                    }
                } catch (exception) {
                    if (exception.message.indexOf("is not a function") >= 0) {
                        const alexaCommand = typeof msg.command === "string" ? msg.command + ":" : "";
                        node.warn("Input '" +  alexaCommand + msg.payload + "' generated command '" +
                            exception.message.match(/[^_a-zA-Z]([_0-9a-zA-Z]*) is not a function/)[1] + "' not supported by device");
                    } else {
                        node.warn(exception);
                    }
                }
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, unitAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-heat-out ") + ": missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    deviceAddress = path.slice(1, 2);
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["thermostat1", "thermostat2", "thermostat3", "thermostat4"]);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-heat-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            parseCommand(protocolName, deviceAddress, msg);
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-heat-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-heat-out", RfxHeaterOutNode);

// An input node for listening to messages from blinds remote controls
    function RfxBlindsInNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;
        this.blinds1Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.blinds1[evt.subtype] || "BLINDS_UNKNOWN") + "/" + evt.id;
            if (evt.subtype !== 2 && evt.subtype !== 4 && evt.subtype !== 5 &&
                evt.subtype !== 10 && evt.subtype !== 18) {
                msg.topic = msg.topic + "/" + evt.unitCode;
            }
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.payload = evt.command;
                node.send(msg);
            }
        };
        this.lighting5Handler = function (evt) {
            let msg = {status: {rssi: evt.rssi}};
            msg.topic = (rfxcom.lighting5[evt.subtype] || "LIGHTING5_UNKNOWN") + "/" + evt.id;
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                if (evt.subtype === rfxcom.lighting5.LIGHTWAVERF) {
                    switch (evt.commandNumber) {
                        case 13:
                        case 14:
                        case 15:
                            msg.payload = evt.command;
                            node.send(msg);
                            break;

                        default:
                            return;
                    }
                }
            }
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    if (node.rfxtrx) {
                        node.rfxtrx.removeListener("blinds1", node.blinds1Handler);
                        node.rfxtrx.removeListener("lighting5", node.lighting5Handler);
                    }
                    releasePort(node);
                });
                node.rfxtrx.on("blinds1", this.blinds1Handler);
                node.rfxtrx.on("lighting5", this.lighting5Handler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-blinds-in", RfxBlindsInNode);

// An output node for sending messages to blind and curtains motors
    function RfxBlindsOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        const node = this;

        node.rfyVenetianMode = node.rfxtrxPort.rfyVenetianMode || "EU";

        // Convert the message payload to the appropriate command, depending on the protocol name (subype)
        const parseCommand = function (protocolName, address, payload) {
            if (/open/i.test(payload) || payload > 0) {
                if (protocolName === "BLINDS_T5") {
                    node.rfxtrx.transmitters[protocolName].down(address);
                } else if (node.rfxtrx.transmitters[protocolName].packetType === "rfy") {
                    node.rfxtrx.transmitters[protocolName].venetianOpen(address);
                } else if (node.rfxtrx.transmitters[protocolName].packetType === "lighting5") {
                    node.rfxtrx.transmitters[protocolName].relayOpen(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].open(address);
                }
            } else if (/close/i.test(payload) || payload < 0) {
                if (protocolName === "BLINDS_T5") {
                    node.rfxtrx.transmitters[protocolName].up(address);
                } else if (node.rfxtrx.transmitters[protocolName].packetType === "rfy") {
                    node.rfxtrx.transmitters[protocolName].venetianClose(address);
                } else if (node.rfxtrx.transmitters[protocolName].packetType === "lighting5") {
                    node.rfxtrx.transmitters[protocolName].relayClose(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].close(address);
                }
            } else if (/stop/i.test(payload) || payload === 0) {
                if (node.rfxtrx.transmitters[protocolName].packetType === "lighting5") {
                    node.rfxtrx.transmitters[protocolName].relayStop(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].stop(address);
                }
            } else if (/confirm|pair|program/i.test(payload)) {
                if (node.rfxtrx.transmitters[protocolName].packetType === "curtain1" ||
                    node.rfxtrx.transmitters[protocolName].packetType === "rfy" ||
                    node.rfxtrx.transmitters[protocolName].packetType === "lighting5") {
                    node.rfxtrx.transmitters[protocolName].program(address);
                } else {
                    node.rfxtrx.transmitters[protocolName].confirm(address);
                }
            } else if (/inter|pos/i.test(payload)) {
                let position = 2;
                const match = /[0-9]+(\.[0-9]*)?/.exec(payload);
                if (match !== null) {
                    position = parseFloat(match[0]);
                }
                if (isNaN(position)) {
                    position = 2;
                } else {
                    node.rfxtrx.transmitters[protocolName].intermediatePosition(address, position);
                }
            } else if (/up/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].up(address);
            } else if (/down/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].down(address);
            } else if (/angle|turn/i.test(payload)) {
                if (/increase|\+/i.test(payload)) {
                    node.rfxtrx.transmitters[protocolName].venetianIncreaseAngle(address);
                } else if (/decrease|-/i.test(payload)) {
                    node.rfxtrx.transmitters[protocolName].venetianDecreaseAngle(address);
                }
            } else if (/auto/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].enableSunSensor(address);
            } else if (/man/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].disableSunSensor(address);
            } else if (/list/i.test(payload)) {
                if (node.rfxtrx.transmitters[protocolName].packetType === "rfy") {
                    node.rfxtrx.transmitters[protocolName].listRemotes(address);
                }
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    let path = [], protocolName, subtype, deviceAddress, unitAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn((node.name || "rfx-blinds-out") + ": missing topic");
                        return;
                    }
                    protocolName = path[0].trim().replace(/ +/g, '_').toUpperCase();
                    if (protocolName === "BLINDS_T2" || protocolName === "BLINDS_T4" ||
                        protocolName === "BLINDS_T5" || protocolName === "BLINDS_T10") {
                        unitAddress = 0;
                        if (path.length > 2) {
                            node.warn((node.name || "rfx-blinds-out") + ": ignoring unit code");
                        }
                    } else if (path.length < 3) {
                        node.warn((node.name || "rfx-blinds-out") + ": missing unit code");
                        return;
                    } else {
                        unitAddress = parseUnitAddress(path[2]);
                        if (protocolName === "BLINDS_T3") {
                            if (unitAddress === 0) {
                                unitAddress = 0x10;
                            } else {
                                unitAddress = unitAddress - 1;
                            }
                        }
                        if (protocolName === "BLINDS_T12") {
                            if (unitAddress === 0) {
                                unitAddress = 0x0f;
                            } else {
                                unitAddress = unitAddress - 1;
                            }
                        }
                    }
                    deviceAddress = path.slice(1, 2);
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName, ["blinds1", "blinds2", "curtain1", "lighting5", "rfy"],
                                                                     {venetianBlindsMode: node.rfyVenetianMode});
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-blinds-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            try {
                                parseCommand(protocolName, deviceAddress.concat(unitAddress), msg.payload);
                            } catch (exception) {
                                if (exception.message.indexOf("is not a function") >= 0) {
                                    const alexaCommand = typeof msg.command === "string" ? msg.command + ":" : "";
                                    node.warn((node.name || "rfx-blinds-out") + ": Input '" + alexaCommand + msg.payload + "' generated command '" +
                                        exception.message.match(/[^_a-zA-Z]([_0-9a-zA-Z]*) is not a function/)[1] + "' not supported by device");
                                } else {
                                    node.warn((node.name || "rfx-blinds-out") + ": " + exception);
                                }
                            }
                        }
                    } catch (exception) {
                        node.warn((node.name || "rfx-blinds-out ") + ": serial port " + node.rfxtrxPort.port + " does not exist");
                    }
                });
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-blinds-out", RfxBlindsOutNode);

};
