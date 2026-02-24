sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "com/incture/announcements/utils/formatter"
], (Controller, JSONModel, MessageToast, MessageBox, formatter) => {
    "use strict";

    return Controller.extend("com.incture.announcements.controller.BaseController", {

        formatter: formatter,

        _navBack: function () {
            this.getOwnerComponent().getRouter().navTo("Announcement");
        },

        onNavBack: function () {
            this._navBack();
        },

        _getCSRFToken: function () {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/",
                    method: "GET",
                    headers: { "X-CSRF-Token": "Fetch" },
                    success: (data, textStatus, request) => {
                        resolve(request.getResponseHeader("X-CSRF-Token"));
                    },
                    error: (xhr, status, err) => {
                        console.error("CSRF token fetch failed:", status, err);
                        reject(err);
                    }
                });
            });
        },

        getCurrentUserEmail: async function () {
            const appId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
            const appModulePath = jQuery.sap.getModulePath(appId.replaceAll(".", "/"));
            const oUserModel = new JSONModel();
            await oUserModel.loadData(appModulePath + "/user-api/currentUser");
            const data = oUserModel.getData();
            if (data && data.email) { return data.email; }
            throw new Error("Email not found in the response.");
        },

        _isPublishLater: function (sStartAnnouncement) {
            if (!sStartAnnouncement) { return false; }
            const oStart = new Date(sStartAnnouncement);
            oStart.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);
            return oStart > oToday;
        },

        _setDefaultDates: function (sModelName) {
            const oModel = this.getView().getModel(sModelName);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            oModel.setProperty("/publishDate", formatter.formatDateToValue(oToday));

            const oExpiry = new Date(oToday);
            oExpiry.setDate(oExpiry.getDate() + 30);
            oModel.setProperty("/expiryDate", formatter.formatDateToValue(oExpiry));

            const oMinExpiry = new Date(oToday);
            oMinExpiry.setDate(oMinExpiry.getDate() + 1);
            oModel.setProperty("/minExpiryDate", oMinExpiry);

            oModel.setProperty("/publishDateEnabled", false);
            oModel.setProperty("/showPublishTodayText", true);
        },

        _onPublishDateChange: function (oEvent, sModelName) {
            const sValue = oEvent.getParameter("value");
            const oModel = this.getView().getModel(sModelName);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            oModel.setProperty("/publishDate", sValue);

            if (!sValue) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", oBundle.getText("publishDateRequired"));
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (oSelected < oToday) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", oBundle.getText("publishDatePastError"));
                return;
            }

            oModel.setProperty("/publishDateValueState", "None");
            oModel.setProperty("/publishDateValueStateText", "");

            const oExpiry = new Date(oSelected);
            oExpiry.setDate(oExpiry.getDate() + 30);
            oModel.setProperty("/expiryDate", formatter.formatDateToValue(oExpiry));
            oModel.setProperty("/expiryDateValueState", "None");
            oModel.setProperty("/expiryDateValueStateText", "");

            const oMinExpiry = new Date(oSelected);
            oMinExpiry.setDate(oMinExpiry.getDate() + 1);
            oModel.setProperty("/minExpiryDate", oMinExpiry);

            this._handleResetButtonVisibility();
        },

        _onExpiryDateChange: function (oEvent, sModelName) {
            const sValue = oEvent.getParameter("value");
            const oModel = this.getView().getModel(sModelName);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            oModel.setProperty("/expiryDate", sValue);

            if (!sValue) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDateRequired"));
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (oSelected < oToday) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDatePastError"));
                return;
            }

            const sPublishDate = oModel.getProperty("/publishDate");
            if (sPublishDate) {
                const oPublish = new Date(sPublishDate);
                oPublish.setHours(0, 0, 0, 0);
                if (oSelected <= oPublish) {
                    oModel.setProperty("/expiryDateValueState", "Error");
                    oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDateBeforePublishError"));
                    return;
                }
            }

            oModel.setProperty("/expiryDateValueState", "None");
            oModel.setProperty("/expiryDateValueStateText", "");
            this._handleResetButtonVisibility();
        },

        _onPublishLaterChange: function (oEvent, sModelName) {
            const bState = oEvent.getParameter("state");
            const oModel = this.getView().getModel(sModelName);

            oModel.setProperty("/publishLater", bState);

            const oBaseDate = new Date();
            if (bState) { oBaseDate.setDate(oBaseDate.getDate() + 1); }
            oBaseDate.setHours(0, 0, 0, 0);

            oModel.setProperty("/publishDate", formatter.formatDateToValue(oBaseDate));
            oModel.setProperty("/publishDateEnabled", bState);
            oModel.setProperty("/showPublishTodayText", !bState);

            const oExpiry = new Date(oBaseDate);
            oExpiry.setDate(oExpiry.getDate() + 30);
            oModel.setProperty("/expiryDate", formatter.formatDateToValue(oExpiry));

            const oMinExpiry = new Date(oBaseDate);
            oMinExpiry.setDate(oMinExpiry.getDate() + 1);
            oModel.setProperty("/minExpiryDate", oMinExpiry);

            this._handleResetButtonVisibility();
        },

        _buildStatusFields: function (sPublishDate, sExpiryDate, bPublishLater, sCurrentDateTime) {
            const sNow = sCurrentDateTime || new Date().toISOString();
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);
            const oPublishDate = new Date(sPublishDate);
            oPublishDate.setHours(0, 0, 0, 0);

            let announcementStatus, startAnnouncement, publishedAt;

            if (!bPublishLater && oPublishDate.getTime() === oToday.getTime()) {
                announcementStatus = "PUBLISHED";
                startAnnouncement = sNow;
                publishedAt = sNow;
            } else {
                announcementStatus = "TO_BE_PUBLISHED";
                startAnnouncement = new Date(sPublishDate).toISOString();
                publishedAt = new Date(sPublishDate).toISOString();
            }

            const oEndDate = new Date(sExpiryDate);
            oEndDate.setDate(oEndDate.getDate() + 1);
            const endAnnouncement = oEndDate.toISOString();

            return { announcementStatus, startAnnouncement, endAnnouncement, publishedAt };
        },

        onResetPress: function () {
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            MessageBox.confirm(oBundle.getText("resetConfirmMessage"), {
                title: oBundle.getText("resetConfirmTitle"),
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.NO,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) { this._performReset(); }
                }
            });
        },

        formatDateOnly: function (oDate) { return formatter.formatDateOnly(oDate); },
        formatUSDateTime: function (oDate) { return formatter.formatUSDateTime(oDate); },
        formatDateToDDMMYYYY: function (oDate) { return formatter.formatDateToDDMMYYYY(oDate); },
        _formatDateToDisplay: function (oDate) { return formatter.formatDateToDisplay(oDate); },
        formatStatusState: function (sStatus) { return formatter.formatStatusState(sStatus); },
        formatStatusText: function (sStatus) { return formatter.formatStatusText(sStatus); }
    });
});