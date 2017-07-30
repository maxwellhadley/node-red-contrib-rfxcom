/*
    Copyright (c) 2014 .. 2017, Maxwell Hadley
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
    var rfxcom = require("rfxcom");

// Set the rfxcom debug option from the environment variable
    var debugOption = {};
    if (process.env.hasOwnProperty("RED_DEBUG") && process.env.RED_DEBUG.indexOf("rfxcom") >= 0) {
        debugOption = {debug: true};
    }

// The config node holding the (serial) port device path for one or more rfxcom family nodes
    function RfxtrxPortNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
    }

// Register the config node
    RED.nodes.registerType("rfxtrx-port", RfxtrxPortNode);

// An object maintaining a pool of config nodes
    var rfxcomPool = function () {
        var pool = {}, intervalTimer = null;

        var connectTo = function (rfxtrx, node) {
            //noinspection JSUnusedLocalSymbols
            rfxtrx.initialise(function (error, response, sequenceNumber) {
                node.log("connected: Serial port " + rfxtrx.device);
                if (intervalTimer !== null) {
                    clearInterval(intervalTimer);
                    intervalTimer = null;
                }
            });
        };

        return {
            get: function (node, port, options) {
                // Returns the RfxCom object associated with port, or creates a new RfxCom object,
                // associates it with the port, and returns it. 'port' is the device file path to
                // the pseudo-serialport, e.g. '/dev/tty.usb-123456'
                var rfxtrx;
                if (!pool[port]) {
                    rfxtrx = new rfxcom.RfxCom(port, options || debugOption);
                    rfxtrx.transmitters = {};
                    rfxtrx.on("connecting", function () {
                        node.log("connecting to " + port);
                        pool[port].references.forEach(function (node) {
                            node.status({fill:"yellow",shape:"dot",text:"connecting..."});
                        });
                    });
                    rfxtrx.on("connectfailed", function (msg) {
                        if (intervalTimer === null) {
                            node.log("connect failed: " + msg);
                            intervalTimer = setInterval(function () {
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
                    // TODO - check if new RFY 'list' (remotes) event handler may be needed?
                    rfxtrx.on("disconnect", function (msg) {
                        node.log("disconnected: " + msg);
                        pool[port].references.forEach(function (node) {
                                showConnectionStatus(node);
                            });
                        if (intervalTimer === null) {
                            intervalTimer = setInterval(function () {
                                connectTo(rfxtrx, node)
                            }, rfxtrx.initialiseWaitTime);
                        }
                    });
                    pool[port] = {rfxtrx: rfxtrx, references: []};
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
                        delete pool[port].rfxtrx;
                        delete pool[port];
                    }
                }
            }
        }
    }();

    var releasePort = function (node) {
        // Decrement the reference count on the node port
        if (node.rfxtrxPort) {
            rfxcomPool.release(node, node.rfxtrxPort.port);
        }
    };

// Utility function: normalises the accepted representations of 'unit addresses' to be
// an integer, and ensures all variants of the 'group address' are converted to 0
    var parseUnitAddress = function (str) {
        if (str === undefined || /all|group|\+/i.test(str)) {
            return 0;
        } else {
            return Math.round(Number(str));
        }
    };

// Flag values for the different types of message packet recognised by the nodes in this file
// By convention, the all-uppercase equivalent of the node-rfxcom object implementing the
// message packet type. The numeric values are the protocol codes used by the RFXCOM API,
// though this is also an arbitrary choice
    var txTypeNumber = {
        LIGHTING1: 0x10,
        LIGHTING2: 0x11,
        LIGHTING3: 0x12,
        LIGHTING4: 0x13,
        LIGHTING5: 0x14,
        LIGHTING6: 0x15,
        CHIME1: 0x16,
        CURTAIN1:  0x18
    };

// This function takes a protocol name and returns the subtype number (defined by the RFXCOM
// API) for that protocol. It also creates the node-rfxcom object implementing the message packet type
// corresponding to that subtype, or re-uses a pre-existing object that implements it. These objects
// are stored in the transmitters property of rfxcomObject
    var getRfxcomSubtype = function (rfxcomObject, protocolName) {
        var subtype;
        if (rfxcomObject.transmitters.hasOwnProperty(protocolName) === false) {
            if ((subtype = rfxcom.lighting1[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting1.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING1
                };
            } else if ((subtype = rfxcom.lighting2[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting2.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING2
                };
            } else if ((subtype = rfxcom.lighting3[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting3.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING3
                };
            } else if ((subtype = rfxcom.lighting4[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting4.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING4
                };
            } else if ((subtype = rfxcom.lighting5[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting5.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING5
                };
            } else if ((subtype = rfxcom.lighting6[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.lighting6.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.LIGHTING6
                };
            } else if ((subtype = rfxcom.chime1[protocolName]) !== undefined) {
                rfxcomObject.transmitters[protocolName] = {
                    tx:   new rfxcom.chime1.transmitter(rfxcomObject, subtype),
                    type: txTypeNumber.CHIME1
                };
            } else {
                subtype = -1; // Error return
                // throw new Error("Protocol type '" + protocolName + "' not supported");
            }
        } else {
            subtype = rfxcomObject.transmitters[protocolName].tx.subtype;
        }
        return subtype;
    };

// Convert a string containing a slash/delimited/path to an Array of (string) parts, removing any empty components
// Return value may be zero-length
    var stringToParts = function (str) {
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
    var normaliseTopic = function (rawTopic) {
        var parts;
        if (rawTopic === undefined || typeof rawTopic !== "string") {
            return [];
        }
        parts = stringToParts(rawTopic);
        if (parts.length >= 1) {
            parts[0] = parts[0].trim().replace(/ +/g, '_').toUpperCase();
        }
        if (parts.length >= 2) {
            // handle housecodes as a special case (X10, ARC, etc)
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
    var normaliseAndCheckTopic = function (topic, pattern) {
        return checkTopic(normaliseTopic(topic), pattern);
    };

// Check if the supplied topic starts with the given pattern (both being normalised)
    var checkTopic = function (parts, pattern) {
        for (var i = 0; i < pattern.length; i++) {
            if (parts[i] !== pattern[i]) {
                return false;
            }
        }
        return true;
    };

// Show the connection status of the node depending on its underlying rfxtrx object
    var showConnectionStatus = function (node) {
        if (node.rfxtrx.connected === false) {
            node.status({fill: "red", shape: "ring", text: "disconnected"});
        } else {
            node.status({fill: "green", shape: "dot",
                text: "OK (v" + node.rfxtrx.firmwareVersion + " " + node.rfxtrx.firmwareType + ")"});
        }
    };

// Purge retransmission timers associated with this node and delete them to avoid leaking memory
    var purgeTimers = function () {
           var tx;
           for (tx in this.retransmissions) {
               if (this.retransmissions.hasOwnProperty(tx)) {
                   this.clearRetransmission(this.retransmissions[tx]);
                   delete this.retransmissions[tx];
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

// An input node for listening to messages from lighting remote controls
    function RfxLightsInNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        this.lighting1Handler = function (evt) {
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.lighting1[evt.subtype] + "/" + evt.housecode;
            if (evt.commandNumber === 5 || evt.commandNumber === 6) {
                msg.topic = msg.topic + "/+";
            } else {
                msg.topic = msg.topic + "/" + evt.unitcode;
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
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.lighting2[evt.subtype] + "/" + evt.id;
            if (evt.commandNumber > 2) {
                msg.topic = msg.topic + "/+";
            } else {
                msg.topic = msg.topic + "/" + evt.unitcode;
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
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.lighting5[evt.subtype] + "/" + evt.id;
            if ((evt.commandNumber === 2 && (evt.subtype === 0 || evt.subtype === 2 || evt.subtype === 4) ) ||
                (evt.commandNumber === 3) && (evt.subtype === 2 || evt.subtype === 4)) {
                msg.topic = msg.topic + "/+";
            } else {
                msg.topic = msg.topic + "/" + evt.unitcode;
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

                    case 2:
                    case 4: // BBSB & CONRAD
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
                }
                node.send(msg);
            }
        };
        this.lighting6Handler = function (evt) {
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.lighting6[evt.subtype] + "/" + evt.id + "/" + evt.groupcode;
            if (evt.commandNumber > 1) {
                msg.topic = msg.topic + "/+";
            } else {
                msg.topic = msg.topic + "/" + evt.unitcode;
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
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.rfxtrx.on("lighting1", this.lighting1Handler);
                node.rfxtrx.on("lighting2", this.lighting2Handler);
                node.rfxtrx.on("lighting5", this.lighting5Handler);
                node.rfxtrx.on("lighting6", this.lighting6Handler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-lights-in", RfxLightsInNode);

// Remove the message event handlers on close
    RfxLightsInNode.prototype.close = function () {
        if (this.rfxtrx) {
            this.rfxtrx.removeListener("lighting1", this.lighting1Handler);
            this.rfxtrx.removeListener("lighting2", this.lighting2Handler);
            this.rfxtrx.removeListener("lighting5", this.lighting5Handler);
            this.rfxtrx.removeListener("lighting6", this.lighting6Handler);
        }
    };

// An input node for listening to messages from PT622 (lighting4) devices
    function RfxPT2262InNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "all";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.devices = RED.nodes.getNode(n.deviceList).devices || [];
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        this.lighting4Handler = function (evt) {
            var msg = {status: {rssi: evt.rssi}};
            var db = node.devices.filter(function (entry) {return entry.rawData === evt.data});
            if (db.length === 0) {
                msg.raw = {data: evt.data, pulseWidth: evt.pulseWidth};
            } else {
                if (node.topicSource === "all" || checkTopic(db[0].device, node.topic)) {
                    msg.topic = db[0].device.join("/");
                    msg.payload = db[0].payload;
                } else {
                    return;
                }
            }
            node.send(msg);
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.rfxtrx.on("lighting4", this.lighting4Handler);
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-PT2262-in", RfxPT2262InNode);

// Remove the message event handlers on close
    RfxPT2262InNode.prototype.close = function () {
        if (this.rfxtrx) {
            this.rfxtrx.removeListener("lighting4", this.lighting4Handler);
        }
    };

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

        var node = this;
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
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                getRfxcomSubtype(node.rfxtrx, "PT2262");
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    var db, topic, lastCommand, rawData = null, pulseWidth = null;
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
                            node.rfxtrx.transmitters['PT2262'].tx.sendData(rawData, pulseWidth);
                            // If we reach this point, the command did not throw an error. Check if should retransmit
                            // it & set the Timeout or Interval as appropriate
                            if (node.retransmit !== "none") {
                                // Wrap the parseCommand arguments, and the retransmission key (=topic) in a
                                // function context, where the function lastCommand() can find them
                                topic = topic.join("/");
                                lastCommand = (function () {
                                    var _rawData = rawData, _pulseWidth = pulseWidth, key = topic;
                                    return function () {
                                        node.rfxtrx.transmitters['PT2262'].tx.sendData(_rawData, _pulseWidth);
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

// Remove all retransmission timers on close
    RfxPT2262OutNode.prototype.close = purgeTimers;

// An input node for listening to messages from (mainly weather) sensors
    function RfxWeatherSensorNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        var i;

        var sendWeatherMessage = function (evt, msg) {
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi};
                if (evt.hasOwnProperty("batteryLevel")) {
                    msg.status.battery = evt.batteryLevel;
                }
                msg.payload = {};
                if (evt.hasOwnProperty("temperature")) {
                    msg.payload.temperature = {value: evt.temperature, unit: "degC"};
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
            sendWeatherMessage(evt, {topic:rfxcom.bbq1[evt.subtype] + "/" + evt.id})
        };
        this.temperaturerainHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.temperatureRain1[evt.subtype] + "/" + evt.id})
        };
        this.temperatureHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.temperature1[evt.subtype] + "/" + evt.id})
        };
        this.humidityHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.humidity1[evt.subtype] + "/" + evt.id})
        };
        this.temperaturehumidityHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.temperatureHumidity1[evt.subtype] + "/" + evt.id})
        };
        this.temphumbaroHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.tempHumBaro1[evt.subtype] + "/" + evt.id})
        };
        this.rainHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.rain1[evt.subtype] + "/" + evt.id})
        };
        this.windHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.wind1[evt.subtype] + "/" + evt.id})
        };
        this.uvHandler = function(evt) {
            sendWeatherMessage(evt, {topic:rfxcom.uv1[evt.subtype] + "/" + evt.id})
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.rfxtrx.on("bbq1", this.bbq1Handler);
                for (i = 1; i < rfxcom.temperatureRain1.length; i++) {
                    node.rfxtrx.on("temprain" + i, this.temperaturerainHandler);
                }
                for (i = 1; i < rfxcom.temperature1.length; i++) {
                    node.rfxtrx.on("temp" + i, this.temperatureHandler);
                }
                for (i = 1; i < rfxcom.humidity1.length; i++) {
                    node.rfxtrx.on("humidity" + i, this.humidityHandler);
                }
                for (i = 1; i < rfxcom.temperatureHumidity1.length; i++) {
                    node.rfxtrx.on("th" + i, this.temperaturehumidityHandler);
                }
                for (i = 1; i < rfxcom.tempHumBaro1.length; i++) {
                    node.rfxtrx.on("thb" + i, this.temphumbaroHandler);
                }
                for (i = 1; i < rfxcom.rain1.length; i++) {
                    node.rfxtrx.on("rain" + i, this.rainHandler);
                }
                for (i = 1; i < rfxcom.wind1.length; i++) {
                    node.rfxtrx.on("wind" + i, this.windHandler);
                }
                for (i = 1; i < rfxcom.uv1.length; i++) {
                    node.rfxtrx.on("uv" + i, this.uvHandler);
                }
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-sensor", RfxWeatherSensorNode);

// Remove the message event handlers on close
    RfxWeatherSensorNode.prototype.close = function () {
        var i;
        if (this.rfxtrx) {
            this.rfxtrx.removeListener("bbq1", this.bbq1Handler);
            for (i = 1; i < rfxcom.temperatureRain1.length; i++) {
                this.rfxtrx.removeListener("temprain" + i, this.temperaturerainHandler);
            }
            for (i = 1; i < rfxcom.temperature1.length; i++) {
                this.rfxtrx.removeListener("temp" + i, this.temperatureHandler);
            }
            for (i = 1; i < rfxcom.humidity1.length; i++) {
                this.rfxtrx.removeListener("humidity" + i, this.humidityHandler);
            }
            for (i = 1; i < rfxcom.temperatureHumidity1.length; i++) {
                this.rfxtrx.removeListener("th" + i, this.temperaturehumidityHandler);
            }
            for (i = 1; i < rfxcom.tempHumBaro1.length; i++) {
                this.rfxtrx.removeListener("thb" + i, this.temphumbaroHandler);
            }
            for (i = 1; i < rfxcom.rain1.length; i++) {
                this.rfxtrx.removeListener("rain" + i, this.rainHandler);
            }
            for (i = 1; i < rfxcom.wind1.length; i++) {
                this.rfxtrx.removeListener("wind" + i, this.windHandler);
            }
            for (i = 1; i < rfxcom.uv1.length; i++) {
                this.rfxtrx.removeListener("uv" + i, this.uvHandler);
            }
        }
    };

// An input node for listening to messages from (electrical) energy & current monitors
    function RfxEnergyMeterNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        var i;

        var sendMeterMessage = function (evt, msg) {
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
                node.send(msg);
            }
        };
        this.elec1Handler = function (evt) {
            sendMeterMessage(evt, {topic: rfxcom.elec1[evt.subtype] + "/" + evt.id})
        };
        this.elec23Handler = function (evt) {
            sendMeterMessage(evt, {topic: rfxcom.elec23[evt.subtype] + "/" + evt.id})
        };
        this.elec4Handler = function (evt) {
            sendMeterMessage(evt, {topic: rfxcom.elec4[evt.subtype] + "/" + evt.id})
        };
        this.elec5Handler = function (evt) {
            sendMeterMessage(evt, {topic: rfxcom.elec5[evt.subtype] + "/" + evt.id})
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                for (i = 1; i < rfxcom.elec1.length; i++) {
                    node.rfxtrx.on("elec" + i, this.elec1Handler)
                }
                for (i = 1; i < rfxcom.elec23.length; i++) {
                    node.rfxtrx.on("elec" + (i + 1), this.elec23Handler)
                }
                for (i = 1; i < rfxcom.elec4.length; i++) {
                    node.rfxtrx.on("elec" + (i + 3), this.elec4Handler)
                }
                for (i = 1; i < rfxcom.elec5.length; i++) {
                    node.rfxtrx.on("elec" + (i + 4), this.elec5Handler)
                }
            }
        } else {
            node.error("missing config: rfxtrx-port");
        }
    }

    RED.nodes.registerType("rfx-meter", RfxEnergyMeterNode);

// Remove the message event handlers on close
    RfxEnergyMeterNode.prototype.close = function () {
		var i;
        if (this.rfxtrx) {
			for (i = 1; i < rfxcom.elec1.length; i++) {
				this.rfxtrx.removeListener("elec" + i, this.elec1Handler);
			}
			for (i = 1; i < rfxcom.elec23.length; i++) {
				this.rfxtrx.removeListener("elec" + (i+1), this.elec23Handler);
			}
			for (i = 1; i < rfxcom.elec4.length; i++) {
				this.rfxtrx.removeListener("elec" + (i+3), this.elec4Handler);
			}
			for (i = 1; i < rfxcom.elec5.length; i++) {
				this.rfxtrx.removeListener("elec" + (i+4), this.elec5Handler);
			}
        }
    };

// An input node for listening to messages from security and smoke detectors
    function RfxDetectorsNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        node.heartbeats = {};
        node.HEARTBEATDELAY = []; // delay in minutes before declaring a detector has gone ailent
        node.HEARTBEATDELAY[rfxcom.security1.X10_DOOR] = 90;
        node.HEARTBEATDELAY[rfxcom.security1.X10_PIR] = 90;
        node.HEARTBEATDELAY[rfxcom.security1.POWERCODE_DOOR] = 20;
        node.HEARTBEATDELAY[rfxcom.security1.POWERCODE_PIR] = 20;
        this.security1Handler = function (evt) {
            var msg = {topic: rfxcom.security1[evt.subtype] + "/" + evt.id};
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                msg.status = {rssi: evt.rssi};
                switch (evt.subtype) {
                    case rfxcom.security1.KD101:
                    case rfxcom.security1.SA30:
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
                    // These detectors send "heartbeat" messages at more or less regular intervals
                    case rfxcom.security1.POWERCODE_DOOR:
                    case rfxcom.security1.POWERCODE_PIR:
                    case rfxcom.security1.X10_DOOR:
                    case rfxcom.security1.X10_PIR:
                        msg.status.battery = evt.batteryLevel;
                        msg.status.tampered = Boolean(evt.tampered);
                        // Clear any existing heartbeat timeout & set a new one
                        if (node.heartbeats.hasOwnProperty(msg.topic)) {
                            clearInterval(node.heartbeats[msg.topic]);
                        }
                        node.heartbeats[msg.topic] = (function () {
                            var heartbeatStoppedMsg = {
                                topic: msg.topic,
                                payload: "Silent",
                                lastMessageStatus: msg.status,
                                lastMessageTimestamp: Date.now(),
                                lastHeardFrom: new Date().toUTCString()
                            };
                            return setInterval(function () {
                                delete heartbeatStoppedMsg._msgid;
                                node.send(heartbeatStoppedMsg);
                            }, 60*1000*node.HEARTBEATDELAY[evt.subtype]);
                        }());
                        node.heartbeats[msg.topic].unref();
                        if (evt.deviceStatus === rfxcom.security.ALARM) {
                            msg.payload = "Alarm";
                        } else if (evt.deviceStatus === rfxcom.security.MOTION) {
                            msg.payload = "Motion";
                        } else if (evt.tampered) {
                            msg.payload = "Tamper";
                        } else if (evt.batteryLevel === 0) {
                            msg.payload = "Battery Low";
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
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.rfxtrx.on("security1", this.security1Handler)
            }
        }
    }

    RED.nodes.registerType("rfx-detector-in", RfxDetectorsNode);

    RfxDetectorsNode.prototype.close = function () {
        this.rfxtrx.removeListener("security1", this.security1Handler);
        for (var heartbeat in this.heartbeats) {
            if (this.heartbeats.hasOwnProperty(heartbeat)) {
                clearInterval(this.heartbeats[heartbeat]);
            }
        }
    };
    
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

        var node = this;
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
        var parseDimLevel = function (payload, alexaCommand, levelRange) {
            var value = NaN;
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
                    } else if (/dim.*-/i.test(payload)) {
                        return "-";
                    } else if (/dim/i.test(payload)) {
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
                value = parseFloat(/[0-9]+(\.[0-9]*)?/.exec(payload)[0]);
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

        // Parses msg.payload looking for lighting command messages (including Alexa-home command messages), calling
        // the corresponding function in the node-rfxcom API to implement it. All parameter checking is delegated to
        // that API. If no valid command is recognised, this does nothing (quietly), but if the transmitter throws an
        // Error this function does not catch it
        var parseCommand = function (protocolName, address, payload, alexaCommand, levelRange) {
            var level, mood;
            if (/dim|bright|level|%|[0-9]\.|\.[0-9]/i.test(payload) || /Percentage/i.test(alexaCommand)) {
                level = parseDimLevel(payload, alexaCommand, levelRange);
                if (isFinite(level)) {
                    node.rfxtrx.transmitters[protocolName].tx.setLevel(address, level);
                } else if (level === '+') {
                    node.rfxtrx.transmitters[protocolName].tx.increaseLevel(address);
                } else if (level === '-') {
                    node.rfxtrx.transmitters[protocolName].tx.decreaseLevel(address);
                }
            } else if (/on/i.test(payload) || payload === 1 || payload === true) {
                node.rfxtrx.transmitters[protocolName].tx.switchOn(address);
            } else if (/off/i.test(payload) || payload === 0 || payload === false) {
                node.rfxtrx.transmitters[protocolName].tx.switchOff(address);
            } else if (/mood/i.test(payload)) {
                mood = parseInt(/([0-9]+)/.exec(payload));
                if (isFinite(mood)) {
                    node.rfxtrx.transmitters[protocolName].tx.setMood(address, mood);
                }
            } else if (/toggle/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].tx.toggleOnOff(address);
            } else if (/program|learn/i.test(payload)) {
                node.rfxtrx.transmitters[protocolName].tx.program(address);
            }
        };

        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    var path = [], protocolName, subtype, deviceAddress, unitAddress, levelRange, lastCommand, topic;
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
                    deviceAddress = path.slice(1, -1);
                    unitAddress = parseUnitAddress(path.slice(-1)[0]);
                    // The subtype is needed because subtypes within the same protocol might have different dim level ranges
                    //noinspection JSUnusedAssignment
                    try {
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-lights-out ") + ": device type '" + protocolName + "' is not supported");
                        } else {
                            switch (node.rfxtrx.transmitters[protocolName].type) {
                                case txTypeNumber.LIGHTING1 :
                                case txTypeNumber.LIGHTING6 :
                                    levelRange = [];
                                    break;

                                case txTypeNumber.LIGHTING2 :
                                    levelRange = [0, 15];
                                    break;

                                case txTypeNumber.LIGHTING3 :
                                    levelRange = [0, 10];
                                    break;

                                case txTypeNumber.LIGHTING5 :
                                    levelRange = [0, 31];
                                    break;
                            }
                            if (levelRange !== undefined) {
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
                                            var protocol = protocolName, address = deviceAddress.concat(unitAddress),
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
                                        var alexaCommand = typeof msg.command === "string" ? msg.command + ":" : "";
                                        node.warn("Input '" +  alexaCommand + msg.payload + "' generated command '" +
                                            exception.message.match(/[^_a-zA-Z]([_0-9a-zA-Z]*) is not a function/)[1] + "' not supported by device");
                                    } else {
                                        node.warn(exception);
                                    }
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

// Remove all retransmission timers on close
    RfxLightsOutNode.prototype.close = purgeTimers;

// An input node for listening to messages from doorbells
    function RfxDoorbellInNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource;
        this.topic = normaliseTopic(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;
        this.lighting1Handler = function (evt) {
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.lighting1[evt.subtype] + "/" + evt.housecode + "/" + evt.unitcode;
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                if (evt.subtype !== 0x01 || evt.commandNumber !== 7) {
                    return;
                }
                node.send(msg);
            }
        };
        this.chime1Handler = function (evt) {
            var msg = {status: {rssi: evt.rssi}};
            msg.topic = rfxcom.chime1[evt.subtype] + "/" + evt.id;
            if (node.topicSource === "all" || normaliseAndCheckTopic(msg.topic, node.topic)) {
                if (evt.subtype === rfxcom.chime1.BYRON_SX) {
                    msg.payload = evt.commandNumber;
                }
                node.send(msg);
            }
        };
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
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

// Remove the message event handlers on close
    RfxDoorbellInNode.prototype.close = function () {
        if (this.rfxtrx) {
            this.rfxtrx.removeListener("lighting1", this.lighting1Handler);
            this.rfxtrx.removeListener("chime1", this.chime1Handler);
        }
    };

// An output node for sending messages to doorbells
    function RfxDoorbellOutNode(n) {
        RED.nodes.createNode(this, n);
        this.port = n.port;
        this.topicSource = n.topicSource || "msg";
        this.topic = stringToParts(n.topic);
        this.name = n.name;
        this.rfxtrxPort = RED.nodes.getNode(this.port);

        var node = this;

        // Generate the chime command depending on the subtype and tone parameter, if any
        var parseCommand = function (protocolName, address, str) {
            var sound = NaN;
            if (str != undefined) {
                sound = parseInt(str);
            }
            try {
                if (protocolName === "BYRON_SX" && !isNaN(sound)) {
                    node.rfxtrx.transmitters[protocolName].tx.chime(address, sound);
                } else {
                    node.rfxtrx.transmitters[protocolName].tx.chime(address);
                }
            } catch (exception) {
                node.warn(exception);
            }
        };
        
        if (node.rfxtrxPort) {
            node.rfxtrx = rfxcomPool.get(node, node.rfxtrxPort.port);
            if (node.rfxtrx !== null) {
                showConnectionStatus(node);
                node.on("close", function () {
                    releasePort(node);
                });
                node.on("input", function (msg) {
                    // Get the device address from the node topic, or the message topic if the node topic is undefined;
                    // parse the device command from the message payload; and send the appropriate command to the address
                    var path = [], protocolName, subtype, deviceAddress, unitAddress;
                    if (node.topicSource === "node" && node.topic !== undefined) {
                        path = node.topic;
                    } else if (msg.topic !== undefined) {
                        path = stringToParts(msg.topic);
                    }
                    if (path.length === 0) {
                        node.warn("rfx-doorbell-out: missing topic");
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
                        subtype = getRfxcomSubtype(node.rfxtrx, protocolName);
                        if (subtype < 0) {
                            node.warn((node.name || "rfx-lights-out ") + ": device type '" + protocolName + "' is not supported");
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

};
