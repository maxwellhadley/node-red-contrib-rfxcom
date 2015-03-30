node-red-contrib-rfxcom
========================

A collection of <a href="http://nodered.org" target="_new">Node-RED</a> nodes to send and receive home automation
commands and data using an
[RFXCOM RFXtrx433E](http://www.rfxcom.com/epages/78165469.sf/en_GB/?ObjectPath=/Shops/78165469/Products/14103)
home automation controller. Also compatible with the older RFXtrx433 transceivers.

Install
-------

Use npm to install this package locally in the Node-RED data directory (by default, `$HOME/.node-red`):

	cd $HOME/.node-red
	npm install node-red-contrib-rfxcom

Alternatively, it can be installed globally:

    npm install -g node-red-contrib-rfxcom

The nodes will be added to the palette the next time node-RED is started.

Nodes included in the package
-----------------------------

**rfx-lights-in** Receives messages from 'lighting' type devices such as remote controls, PIR sensors used as light
switches, and some types of doorbell.

**rfx-lights-out** Sends messages to 'lighting' type devices including switches, dimmers, and some types of relay.

**rfx-sensor** Receives messages from temperature, humidity, pressure and other weather sensors.

**rfx-meter** Receives messages from wireless energy monitors such as OWL

Basic help text is provided for each node. Additional information is available in the 'RFXmngr.exe' program supplied
with the RFXtrx433E, and more details may be found in the SDK documentation, available on request from RFXCOM.

Example flow
------------

The following Node-RED flow listens to an Oregon temperature sensor, and turns a HomeEasy relay on if the temperature is
9.5 degrees or less, off otherwise:

    [{"id":"dc1031e4.23efd","type":"rfxtrx-port","port":"/dev/ttyUSB0"},{"id":"f858dd6.f07a72","type":"rfx-sensor","name":"","port":"dc1031e4.23efd","topicSource":"single","topic":"TH1/0x8E01","x":113,"y":118,"z":"4235a364.bdca5c","wires":[["66729a21.998d64"]]},{"id":"1b5de8f1.e4a217","type":"rfx-lights-out","name":"","port":"dc1031e4.23efd","topicSource":"node","topic":"AC/0x001EF1CE/4","x":591,"y":215,"z":"4235a364.bdca5c","wires":[]},{"id":"66729a21.998d64","type":"switch","name":"","property":"payload.temperature.value","rules":[{"t":"gt","v":"9.5"},{"t":"else"}],"checkall":"true","outputs":2,"x":226,"y":208,"z":"4235a364.bdca5c","wires":[["41141c7b.beebe4"],["8ac73565.7538c8"]]},{"id":"41141c7b.beebe4","type":"change","action":"replace","property":"payload","from":"","to":"\"Off\"","reg":false,"name":"Turn off","x":396,"y":155,"z":"4235a364.bdca5c","wires":[["1b5de8f1.e4a217"]]},{"id":"8ac73565.7538c8","type":"change","action":"replace","property":"payload","from":"","to":"\"On\"","reg":false,"name":"Turn on","x":397,"y":271,"z":"4235a364.bdca5c","wires":[["1b5de8f1.e4a217"]]}]
