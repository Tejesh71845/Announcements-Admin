sap.ui.define([
    "com/incture/announcements/controller/BaseController",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], (BaseController, JSONModel, MessageToast, MessageBox) => {
    "use strict";

    const BANNER_ANNOUNCEMENT_TYPE = "Banner";
    const MODEL = "bannerModel";

    return BaseController.extend("com.incture.announcements.controller.CreateBannerAnnouncement", {

        onInit: function () {
            this._router = this.getOwnerComponent().getRouter();
            this._router
                .getRoute("CreateBannerAnnouncement")
                .attachPatternMatched(this._onRouteMatched, this);

            this._initBannerModel();
        },

        _onRouteMatched: function (oEvent) {
            const oArgs = oEvent.getParameter("arguments");
            const sEditId = oArgs?.announcementId ?? null;
            const bDuplicate = oArgs?.duplicate === "true";

            if (sEditId && bDuplicate) {
                this._loadAnnouncementForDuplicate(sEditId);
            } else if (sEditId) {
                this._loadAnnouncementForEdit(sEditId);
            } else {
                this._initBannerModel();
                this._setDefaultDates(MODEL);
            }
        },

        _initBannerModel: function () {
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            this.getView().setModel(new JSONModel({
                announcementContent: "",
                contentCharCount: "0/100",

                publishDate: "",
                expiryDate: "",
                publishLater: false,
                publishDateEnabled: false,
                showPublishTodayText: true,
                minPublishDate: oToday,
                minExpiryDate: oToday,

                announcementContentValueState: "None",
                announcementContentValueStateText: "",
                publishDateValueState: "None",
                publishDateValueStateText: "",
                expiryDateValueState: "None",
                expiryDateValueStateText: "",

                resetButtonEnabled: false,
                isEditMode: false,
                isDuplicateMode: false,
                editId: null
            }), MODEL);
        },

        _loadAnnouncementForEdit: function (sAnnouncementId) {
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const oBusy = new sap.m.BusyDialog({ text: oBundle.getText("loadingAnnouncement") });
            oBusy.open();

            this.getOwnerComponent().getModel("announcementModel")
                .read(`/Announcements('${sAnnouncementId}')`, {
                    success: (oData) => { oBusy.close(); this._populateFormForEdit(oData); },
                    error: () => {
                        oBusy.close();
                        MessageBox.error(oBundle.getText("loadAnnouncementError"));
                        this._navBack();
                    }
                });
        },

        _populateFormForEdit: function (oData) {
            const oModel = this.getView().getModel(MODEL);
            const sPublishDate = oData.startAnnouncement
                ? this.formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "";
            const sExpiryDate = oData.endAnnouncement
                ? this.formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "";
            const bPublishLater = this._isPublishLater(oData.startAnnouncement);
            const sContent = oData.title || "";

            oModel.setData({
                announcementContent: sContent,
                contentCharCount: `${sContent.length}/100`,

                publishDate: sPublishDate,
                expiryDate: sExpiryDate,
                publishLater: bPublishLater,
                publishDateEnabled: bPublishLater,
                showPublishTodayText: !bPublishLater,
                minPublishDate: new Date(),
                minExpiryDate: new Date(),

                announcementContentValueState: "None", announcementContentValueStateText: "",
                publishDateValueState: "None", publishDateValueStateText: "",
                expiryDateValueState: "None", expiryDateValueStateText: "",

                resetButtonEnabled: false,
                isEditMode: true,
                isDuplicateMode: false,
                editId: oData.announcementId,

                originalAnnouncementContent: sContent,
                originalPublishDate: sPublishDate,
                originalExpiryDate: sExpiryDate,
                originalPublishLater: bPublishLater
            });
        },

        /* ========================================================
         * LOAD FOR DUPLICATE
         * ======================================================== */

        _loadAnnouncementForDuplicate: function (sAnnouncementId) {
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const oBusy = new sap.m.BusyDialog({ text: oBundle.getText("loadingAnnouncement") });
            oBusy.open();

            this.getOwnerComponent().getModel("announcementModel")
                .read(`/Announcements('${sAnnouncementId}')`, {
                    success: (oData) => { oBusy.close(); this._populateFormForDuplicate(oData); },
                    error: () => {
                        oBusy.close();
                        MessageBox.error(oBundle.getText("loadAnnouncementError"));
                        this._navBack();
                    }
                });
        },

        _populateFormForDuplicate: function (oData) {
            const oModel = this.getView().getModel(MODEL);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            const sContent = oData.title || "";
            const bPublishLater = this._isPublishLater(oData.startAnnouncement);

            let sPublishDate = "", sExpiryDate = "";
            let bPublishDateEnabled = false, bShowPublishTodayText = true;
            let oMinPublishDate = new Date(oToday), oMinExpiryDate = new Date(oToday);

            if (!bPublishLater) {
                sPublishDate = this.formatter.formatDateToValue(oToday);
                oMinExpiryDate = new Date(oToday);
                oMinExpiryDate.setDate(oMinExpiryDate.getDate() + 1);
            } else {
                bPublishDateEnabled = true;
                bShowPublishTodayText = false;
                oMinPublishDate = new Date(oToday);
                oMinPublishDate.setDate(oMinPublishDate.getDate() + 1);
            }

            oModel.setData({
                announcementContent: sContent,
                contentCharCount: `${sContent.length}/100`,

                publishDate: sPublishDate,
                expiryDate: sExpiryDate,
                publishLater: bPublishLater,
                publishDateEnabled: bPublishDateEnabled,
                showPublishTodayText: bShowPublishTodayText,
                minPublishDate: oMinPublishDate,
                minExpiryDate: oMinExpiryDate,

                announcementContentValueState: "None", announcementContentValueStateText: "",
                publishDateValueState: "None", publishDateValueStateText: "",
                expiryDateValueState: "None", expiryDateValueStateText: "",

                isEditMode: false,
                isDuplicateMode: true,
                editId: null,
                resetButtonEnabled: true
            });
        },

        onContentLiveChange: function (oEvent) {
            const sValue = oEvent.getParameter("value") || "";
            this.getView().getModel(MODEL).setProperty("/contentCharCount", `${sValue.length}/100`);
        },

        onInputChange: function (oEvent) {
            const oSource = oEvent.getSource();
            let sValue = oEvent.getParameter("value") || "";

            sValue = sValue.replace(/[^a-zA-Z0-9 ]/g, "");
            const iMax = oSource.getMaxLength?.() > 0 ? oSource.getMaxLength() : 100;
            if (sValue.length > iMax) { sValue = sValue.substring(0, iMax); }
            if (sValue !== oEvent.getParameter("value")) { oSource.setValue(sValue); }

            this.getView().getModel(MODEL).setProperty("/contentCharCount", `${sValue.length}/100`);
            this._handleResetButtonVisibility();
            this._validateField(oSource);
        },

        _validateField: function (oSource) {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (oSource.getId().includes("idBannerAnnouncementContentFld")) {
                const bValid = (oModel.getProperty("/announcementContent") || "").trim().length > 0;
                oModel.setProperty("/announcementContentValueState", bValid ? "None" : "Error");
                oModel.setProperty("/announcementContentValueStateText", bValid ? "" : oBundle.getText("announcementContentRequired"));
            }
        },

        onPublishDateChange: function (oEvent) { this._onPublishDateChange(oEvent, MODEL); },
        onExpiryDateChange: function (oEvent) { this._onExpiryDateChange(oEvent, MODEL); },
        onPublishLaterChange: function (oEvent) { this._onPublishLaterChange(oEvent, MODEL); },


        _handleResetButtonVisibility: function () {
            const oModel = this.getView().getModel(MODEL);

            if (oModel.getProperty("/isEditMode")) {
                const bChanged =
                    oModel.getProperty("/announcementContent") !== oModel.getProperty("/originalAnnouncementContent") ||
                    oModel.getProperty("/publishDate") !== oModel.getProperty("/originalPublishDate") ||
                    oModel.getProperty("/expiryDate") !== oModel.getProperty("/originalExpiryDate") ||
                    oModel.getProperty("/publishLater") !== oModel.getProperty("/originalPublishLater");
                oModel.setProperty("/resetButtonEnabled", bChanged);

            } else if (oModel.getProperty("/isDuplicateMode")) {
                oModel.setProperty("/resetButtonEnabled", true);

            } else {
                const bHasData =
                    (oModel.getProperty("/announcementContent") || "").length > 0 ||
                    (oModel.getProperty("/publishDate") || "").length > 0 ||
                    (oModel.getProperty("/expiryDate") || "").length > 0;
                oModel.setProperty("/resetButtonEnabled", bHasData);
            }
        },

        _validateAllFields: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sContent = (oModel.getProperty("/announcementContent") || "").trim();
            let bValid = true;

            if (!sContent) {
                oModel.setProperty("/announcementContentValueState", "Error");
                oModel.setProperty("/announcementContentValueStateText", oBundle.getText("announcementContentRequired"));
                bValid = false;
            } else {
                oModel.setProperty("/announcementContentValueState", "None");
                oModel.setProperty("/announcementContentValueStateText", "");
            }

            if (!oModel.getProperty("/publishDate")) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", oBundle.getText("publishDateRequired"));
                bValid = false;
            }

            if (!oModel.getProperty("/expiryDate")) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDateRequired"));
                bValid = false;
            }

            return bValid;
        },

        onSubmitPress: function () {
            if (!this._validateAllFields()) {
                MessageToast.show(this.getView().getModel("i18n").getResourceBundle().getText("validationError"));
                return;
            }

            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const bIsEdit = oModel.getProperty("/isEditMode");
            const bIsDup = oModel.getProperty("/isDuplicateMode");

            const { announcementStatus } = this._buildStatusFields(
                oModel.getProperty("/publishDate"),
                oModel.getProperty("/expiryDate"),
                oModel.getProperty("/publishLater")
            );
            const bWillBePublished = announcementStatus === "PUBLISHED";

            if (bWillBePublished) {
                const oAnnouncementModel = this.getOwnerComponent().getModel("announcementModel");

                oAnnouncementModel.read("/Announcements", {
                    filters: [
                        new sap.ui.model.Filter("announcementStatus", sap.ui.model.FilterOperator.EQ, "PUBLISHED"),
                        new sap.ui.model.Filter("announcementType", sap.ui.model.FilterOperator.EQ, "Banner"),
                        new sap.ui.model.Filter("isActive", sap.ui.model.FilterOperator.EQ, true)
                    ],
                    urlParameters: { $select: "announcementId" },
                    success: (oData) => {
                        const aPublished = oData?.results || [];
                        const sCurrentEditId = oModel.getProperty("/editId");

                        const nRelevantCount = bIsEdit
                            ? aPublished.filter(o => o.announcementId !== sCurrentEditId).length
                            : aPublished.length;

                        if (nRelevantCount >= 2) {
                            MessageBox.error(oBundle.getText("maxPublishedBannerError"));
                            return;
                        }

                        this._confirmAndSubmit(bIsEdit, bIsDup, oModel, oBundle);
                    },
                    error: () => {
                        MessageBox.error(oBundle.getText("loadAnnouncementError"));
                    }
                });
            } else {
                this._confirmAndSubmit(bIsEdit, bIsDup, oModel, oBundle);
            }
        },

        _confirmAndSubmit: function (bIsEdit, bIsDup, oModel, oBundle) {
            const sConfirmMsg = oBundle.getText(bIsEdit ? "updateConfirmMessage" : bIsDup ? "duplicateConfirmMessage" : "submitConfirmMessage");
            const sConfirmTitle = oBundle.getText(bIsEdit ? "updateConfirmTitle" : bIsDup ? "duplicateConfirmTitle" : "submitConfirmTitle");

            MessageBox.confirm(sConfirmMsg, {
                title: sConfirmTitle,
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.YES,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        bIsEdit ? this._handleUpdate() : this._handleSubmit();
                    }
                }
            });
        },

        _handleSubmit: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const bIsDup = oModel.getProperty("/isDuplicateMode");

            const oBusy = new sap.m.BusyDialog({
                text: oBundle.getText(bIsDup ? "duplicatingAnnouncement" : "submittingAnnouncement")
            });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    const { announcementStatus, startAnnouncement, endAnnouncement, publishedAt } =
                        this._buildStatusFields(
                            oModel.getProperty("/publishDate"),
                            oModel.getProperty("/expiryDate"),
                            oModel.getProperty("/publishLater")
                        );

                    const oPayload = {
                        data: [{
                            title: (oModel.getProperty("/announcementContent") || "").trim(),
                            description: "Banner Announcement",
                            announcementType: BANNER_ANNOUNCEMENT_TYPE,
                            announcementStatus,
                            startAnnouncement,
                            endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt,
                            category: "General"
                        }]
                    };

                    return this._getCSRFToken().then((csrfToken) =>
                        new Promise((resolve, reject) => {
                            $.ajax({
                                url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/bulkCreateAnnouncements",
                                method: "POST",
                                contentType: "application/json",
                                dataType: "json",
                                headers: { "X-CSRF-Token": csrfToken },
                                data: JSON.stringify(oPayload),
                                success: resolve,
                                error: reject
                            });
                        })
                    ).then(() => {
                        oBusy.close();
                        const sKey = bIsDup
                            ? (announcementStatus === "PUBLISHED" ? "duplicateBannerSuccess" : "duplicateBannerScheduledSuccess")
                            : (announcementStatus === "PUBLISHED" ? "createBannerSuccess" : "createBannerScheduledSuccess");
                        MessageToast.show(oBundle.getText(sKey));
                        this._navBack();
                    });
                })
                .catch((oErr) => {
                    oBusy.close();
                    MessageBox.error(oErr?.responseJSON?.error?.message || oBundle.getText("createAnnouncementError"));
                });
        },

        _handleUpdate: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sNow = new Date().toISOString();

            const oBusy = new sap.m.BusyDialog({ text: oBundle.getText("updatingAnnouncement") });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    const { announcementStatus, startAnnouncement, endAnnouncement, publishedAt } =
                        this._buildStatusFields(
                            oModel.getProperty("/publishDate"),
                            oModel.getProperty("/expiryDate"),
                            oModel.getProperty("/publishLater"),
                            sNow
                        );

                    const oPayload = {
                        title: (oModel.getProperty("/announcementContent") || "").trim(),
                        description: "Banner Announcement",
                        announcementType: BANNER_ANNOUNCEMENT_TYPE,
                        announcementStatus,
                        startAnnouncement,
                        endAnnouncement,
                        publishedBy: sUserEmail,
                        publishedAt,
                        modifiedAt: sNow,
                        modifiedBy: sUserEmail,
                        category: "General"
                    };

                    return this._getCSRFToken().then((csrfToken) =>
                        new Promise((resolve, reject) => {
                            $.ajax({
                                url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${oModel.getProperty("/editId")}')`,
                                method: "PATCH",
                                contentType: "application/json",
                                dataType: "json",
                                headers: { "X-CSRF-Token": csrfToken },
                                data: JSON.stringify(oPayload),
                                success: resolve,
                                error: reject
                            });
                        })
                    ).then(() => {
                        oBusy.close();
                        MessageToast.show(oBundle.getText(
                            announcementStatus === "PUBLISHED" ? "updateBannerSuccess" : "updateBannerScheduledSuccess"
                        ));
                        this._navBack();
                    });
                })
                .catch((oErr) => {
                    oBusy.close();
                    MessageBox.error(oErr?.responseJSON?.error?.message || oBundle.getText("updateAnnouncementError"));
                });
        },

        _performReset: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (oModel.getProperty("/isEditMode")) {
                const sOrig = oModel.getProperty("/originalAnnouncementContent");
                oModel.setProperty("/announcementContent", sOrig);
                oModel.setProperty("/contentCharCount", `${sOrig.length}/100`);
                oModel.setProperty("/publishDate", oModel.getProperty("/originalPublishDate"));
                oModel.setProperty("/expiryDate", oModel.getProperty("/originalExpiryDate"));
                oModel.setProperty("/publishLater", oModel.getProperty("/originalPublishLater"));
                MessageToast.show(oBundle.getText("resetToOriginalMessage"));

            } else if (oModel.getProperty("/isDuplicateMode")) {
                this._initBannerModel();
                this._setDefaultDates(MODEL);
                oModel.setProperty("/isDuplicateMode", true);
                MessageToast.show(oBundle.getText("formReset"));

            } else {
                this._initBannerModel();
                this._setDefaultDates(MODEL);
                MessageToast.show(oBundle.getText("formReset"));
            }

            oModel.setProperty("/resetButtonEnabled", false);
        }
    });
});