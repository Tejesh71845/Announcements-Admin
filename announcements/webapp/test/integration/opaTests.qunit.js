/* global QUnit */
QUnit.config.autostart = false;

sap.ui.require(["com/incture/announcements/test/integration/AllJourneys"
], function () {
	QUnit.start();
});
