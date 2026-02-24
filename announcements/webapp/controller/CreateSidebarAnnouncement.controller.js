sap.ui.define([
    "com/incture/announcements/controller/BaseController",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/richtexteditor/RichTextEditor",
    "sap/ui/richtexteditor/library"
], (BaseController, JSONModel, MessageToast, MessageBox, RTE, library) => {
    "use strict";

    const DEFAULT_CATEGORY = "General";

    const ANNOUNCEMENT_TYPE = {
        SIDEBAR_POPUP: "Sidebar (Popup)",
        SIDEBAR: "Sidebar"
    };

    const MODEL = "sidebarModel";

    return BaseController.extend("com.incture.announcements.controller.CreateSidebarAnnouncement", {


        onInit: function () {
            this._router = this.getOwnerComponent().getRouter();
            this._router
                .getRoute("CreateSidebarAnnouncement")
                .attachPatternMatched(this._onRouteMatched, this);

            this._initSidebarModel();
            this._initCategoryModel();
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
                this._initSidebarModel();
                this._setDefaultDates(MODEL);
                setTimeout(() => { this._initRichTextEditor("idSidebarRichTextCntr"); }, 200);
            }
        },

        _initSidebarModel: function () {
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            this.getView().setModel(new JSONModel({
                title: "",
                titleCharCount: "0/100",
                categories: [],
                description: "",
                descriptionCharCount: "0/500",
                popupAnnouncement: false,

                publishDate: "",
                expiryDate: "",
                publishLater: false,
                publishDateEnabled: false,
                showPublishTodayText: true,
                minPublishDate: oToday,
                minExpiryDate: oToday,

                organizationFinance: true,
                organizationNonFinance: true,
                sectorCorporate: true,
                sectorInnovativeMedicine: true,
                sectorMedTech: true,

                titleValueState: "None",
                titleValueStateText: "",
                descriptionValueState: "None",
                descriptionValueStateText: "",
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

        _initCategoryModel: function () {
            const oCategoryModel = new JSONModel();
            const appId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
            const appModulePath = jQuery.sap.getModulePath(appId.replaceAll(".", "/"));

            oCategoryModel.loadData(appModulePath + "/model/Category.json").catch((oErr) => {
                console.error("Failed to load Category.json:", oErr);
                oCategoryModel.setData({ category: [] });
            });

            this.getView().setModel(oCategoryModel, "categoryModel");
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
            const bIsPopup = oData.announcementType === ANNOUNCEMENT_TYPE.SIDEBAR_POPUP;
            const aCategories = this._normaliseCategoryFromBackend(oData.category);

            const sPublishDate = oData.startAnnouncement
                ? this.formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "";
            const sExpiryDate = oData.endAnnouncement
                ? this.formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "";

            const bPublishLater = this._isPublishLater(oData.startAnnouncement);
            const sTitle = oData.title || "";
            const sDescription = oData.description || "";

            oModel.setData({
                title: sTitle,
                titleCharCount: `${sTitle.length}/100`,
                categories: aCategories,
                description: sDescription,
                descriptionCharCount: `${sDescription.replace(/<[^>]*>/g, "").trim().length}/500`,
                popupAnnouncement: bIsPopup,

                publishDate: sPublishDate,
                expiryDate: sExpiryDate,
                publishLater: bPublishLater,
                publishDateEnabled: bPublishLater,
                showPublishTodayText: !bPublishLater,
                minPublishDate: new Date(),
                minExpiryDate: new Date(),

                organizationFinance: true,
                organizationNonFinance: true,
                sectorCorporate: true,
                sectorInnovativeMedicine: true,
                sectorMedTech: true,

                titleValueState: "None", titleValueStateText: "",
                descriptionValueState: "None", descriptionValueStateText: "",
                publishDateValueState: "None", publishDateValueStateText: "",
                expiryDateValueState: "None", expiryDateValueStateText: "",

                resetButtonEnabled: false,
                isEditMode: true,
                isDuplicateMode: false,
                editId: oData.announcementId,

                originalTitle: sTitle,
                originalCategories: aCategories.slice(),
                originalDescription: sDescription,
                originalPopupAnnouncement: bIsPopup,
                originalPublishDate: sPublishDate,
                originalExpiryDate: sExpiryDate,
                originalPublishLater: bPublishLater
            });

            setTimeout(() => { this._initRichTextEditor("idSidebarRichTextCntr"); }, 200);
        },

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

            const bIsPopup = oData.announcementType === ANNOUNCEMENT_TYPE.SIDEBAR_POPUP;
            const aCategories = this._normaliseCategoryFromBackend(oData.category);
            const sTitle = oData.title || "";
            const sDescription = oData.description || "";
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
                title: sTitle,
                titleCharCount: `${sTitle.length}/100`,
                categories: aCategories,
                description: sDescription,
                descriptionCharCount: `${sDescription.replace(/<[^>]*>/g, "").trim().length}/500`,
                popupAnnouncement: bIsPopup,

                publishDate: sPublishDate,
                expiryDate: sExpiryDate,
                publishLater: bPublishLater,
                publishDateEnabled: bPublishDateEnabled,
                showPublishTodayText: bShowPublishTodayText,
                minPublishDate: oMinPublishDate,
                minExpiryDate: oMinExpiryDate,

                organizationFinance: true,
                organizationNonFinance: true,
                sectorCorporate: true,
                sectorInnovativeMedicine: true,
                sectorMedTech: true,

                titleValueState: "None", titleValueStateText: "",
                descriptionValueState: "None", descriptionValueStateText: "",
                publishDateValueState: "None", publishDateValueStateText: "",
                expiryDateValueState: "None", expiryDateValueStateText: "",

                isEditMode: false,
                isDuplicateMode: true,
                editId: null,
                resetButtonEnabled: true
            });

            setTimeout(() => { this._initRichTextEditor("idSidebarRichTextCntr"); }, 200);
        },

        _normaliseCategoryFromBackend: function (vCategory) {
            if (!vCategory) { return [DEFAULT_CATEGORY]; }
            if (Array.isArray(vCategory)) {
                return vCategory.length > 0 ? vCategory : [DEFAULT_CATEGORY];
            }
            if (typeof vCategory === "string" && vCategory.trim().length > 0) {
                return vCategory.split(",").map(s => s.trim()).filter(Boolean);
            }
            return [DEFAULT_CATEGORY];
        },

        _initRichTextEditor: function (sContainerId) {
            if (this._oRichTextEditor) { this._oRichTextEditor.destroy(); }

            const oModel = this.getView().getModel(MODEL);
            const sDescription = oModel.getProperty("/description") || "";

            this._oRichTextEditor = new RTE({
                editorType: library.EditorType.TinyMCE7,
                width: "100%",
                height: "300px",
                customToolbar: true,
                showGroupFont: true,
                showGroupLink: true,
                showGroupInsert: false,
                value: sDescription,
                ready: function () {
                    this.addButtonGroup("styles").addButtonGroup("table");
                },
                change: (oEvent) => {
                    const sValue = oEvent.getParameter("newValue");
                    const sPlainText = sValue.replace(/<[^>]*>/g, "").trim();
                    const MAX_CHARS = 500;

                    if (sPlainText.length > MAX_CHARS) {
                        if (!this._isCharLimitToastShown) {
                            MessageToast.show(
                                this.getView().getModel("i18n").getResourceBundle()
                                    .getText("descriptionCharLimitError", [MAX_CHARS, sPlainText.length])
                            );
                            this._isCharLimitToastShown = true;
                            setTimeout(() => { this._isCharLimitToastShown = false; }, 3000);
                        }
                        this._oRichTextEditor.setValue(oModel.getProperty("/description") || "");
                        return;
                    }

                    this._isCharLimitToastShown = false;
                    oModel.setProperty("/description", sValue);
                    oModel.setProperty("/descriptionCharCount", `${sPlainText.length}/500`);
                    this._handleResetButtonVisibility();
                    this._validateRichTextDescription();
                }
            });

            const oContainer = this.byId(sContainerId);
            if (oContainer) {
                oContainer.removeAllItems();
                oContainer.addItem(this._oRichTextEditor);
            }
        },

        _validateRichTextDescription: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sPlain = (oModel.getProperty("/description") || "").replace(/<[^>]*>/g, "").trim();
            const bValid = sPlain.length > 0;

            oModel.setProperty("/descriptionValueState", bValid ? "None" : "Error");
            oModel.setProperty("/descriptionValueStateText", bValid ? "" : oBundle.getText("descriptionRequired"));

            const oContainer = this.byId("idSidebarRichTextCntr");
            if (oContainer) {
                bValid
                    ? oContainer.removeStyleClass("richTextErrorCss")
                    : oContainer.addStyleClass("richTextErrorCss");
            }
        },

        onTitleLiveChange: function (oEvent) {
            const sValue = oEvent.getParameter("value") || "";
            this.getView().getModel(MODEL).setProperty("/titleCharCount", `${sValue.length}/100`);
        },

        onInputChange: function (oEvent) {
            const oSource = oEvent.getSource();
            let sValue = oEvent.getParameter("value") || "";

            sValue = sValue.replace(/[^a-zA-Z0-9 ]/g, "");
            const iMax = oSource.getMaxLength?.() > 0 ? oSource.getMaxLength() : 100;
            if (sValue.length > iMax) { sValue = sValue.substring(0, iMax); }
            if (sValue !== oEvent.getParameter("value")) { oSource.setValue(sValue); }

            this.getView().getModel(MODEL).setProperty("/titleCharCount", `${sValue.length}/100`);
            this._handleResetButtonVisibility();
            this._validateField(oSource);
        },

        _validateField: function (oSource) {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (oSource.getId().includes("idSidebarTitleFld")) {
                const bValid = (oModel.getProperty("/title") || "").trim().length > 0;
                oModel.setProperty("/titleValueState", bValid ? "None" : "Error");
                oModel.setProperty("/titleValueStateText", bValid ? "" : oBundle.getText("titleRequired"));
            }
        },

        onCategoryChange: function (oEvent) {
            this.getView().getModel(MODEL).setProperty("/categories", oEvent.getSource().getSelectedKeys());
            this._handleResetButtonVisibility();
        },

        onPopupSwitchChange: function (oEvent) {
            this.getView().getModel(MODEL).setProperty("/popupAnnouncement", oEvent.getParameter("state"));
            this._handleResetButtonVisibility();
        },


        onPublishDateChange: function (oEvent) { this._onPublishDateChange(oEvent, MODEL); },
        onExpiryDateChange: function (oEvent) { this._onExpiryDateChange(oEvent, MODEL); },
        onPublishLaterChange: function (oEvent) { this._onPublishLaterChange(oEvent, MODEL); },

        onOrganizationChange: function () { this._handleResetButtonVisibility(); },
        onSectorChange: function () { this._handleResetButtonVisibility(); },

        _handleResetButtonVisibility: function () {
            const oModel = this.getView().getModel(MODEL);

            if (oModel.getProperty("/isEditMode")) {
                const aCategories = oModel.getProperty("/categories") || [];
                const aOrigCat = oModel.getProperty("/originalCategories") || [];
                const bCatChanged = aCategories.length !== aOrigCat.length ||
                    !aCategories.every(k => aOrigCat.includes(k));

                const bChanged =
                    oModel.getProperty("/title") !== oModel.getProperty("/originalTitle") ||
                    oModel.getProperty("/description") !== oModel.getProperty("/originalDescription") ||
                    bCatChanged ||
                    oModel.getProperty("/popupAnnouncement") !== oModel.getProperty("/originalPopupAnnouncement") ||
                    oModel.getProperty("/publishDate") !== oModel.getProperty("/originalPublishDate") ||
                    oModel.getProperty("/expiryDate") !== oModel.getProperty("/originalExpiryDate") ||
                    oModel.getProperty("/publishLater") !== oModel.getProperty("/originalPublishLater");

                oModel.setProperty("/resetButtonEnabled", bChanged);

            } else if (oModel.getProperty("/isDuplicateMode")) {
                oModel.setProperty("/resetButtonEnabled", true);

            } else {
                const bHasData =
                    (oModel.getProperty("/title") || "").length > 0 ||
                    (oModel.getProperty("/description") || "").length > 0 ||
                    (oModel.getProperty("/publishDate") || "").length > 0 ||
                    (oModel.getProperty("/expiryDate") || "").length > 0;
                oModel.setProperty("/resetButtonEnabled", bHasData);
            }
        },

        _validateAllFields: function () {
            const oModel = this.getView().getModel(MODEL);
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const sPlain = (oModel.getProperty("/description") || "").replace(/<[^>]*>/g, "").trim();
            let bValid = true;

            if (!sTitle) {
                oModel.setProperty("/titleValueState", "Error");
                oModel.setProperty("/titleValueStateText", oBundle.getText("titleRequired"));
                bValid = false;
            } else {
                oModel.setProperty("/titleValueState", "None");
                oModel.setProperty("/titleValueStateText", "");
            }

            if (!sPlain) {
                oModel.setProperty("/descriptionValueState", "Error");
                oModel.setProperty("/descriptionValueStateText", oBundle.getText("descriptionRequired"));
                const oContainer = this.byId("idSidebarRichTextCntr");
                if (oContainer) { oContainer.addStyleClass("richTextErrorCss"); }
                bValid = false;
            } else {
                oModel.setProperty("/descriptionValueState", "None");
                oModel.setProperty("/descriptionValueStateText", "");
                const oContainer = this.byId("idSidebarRichTextCntr");
                if (oContainer) { oContainer.removeStyleClass("richTextErrorCss"); }
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

        _buildCategoryPayload: function () {
            const aCategories = this.getView().getModel(MODEL).getProperty("/categories") || [];
            return aCategories.length > 0 ? aCategories.join(", ") : DEFAULT_CATEGORY;
        },

        _buildAnnouncementType: function () {
            return this.getView().getModel(MODEL).getProperty("/popupAnnouncement")
                ? ANNOUNCEMENT_TYPE.SIDEBAR_POPUP
                : ANNOUNCEMENT_TYPE.SIDEBAR;
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
                            title: (oModel.getProperty("/title") || "").trim(),
                            description: (oModel.getProperty("/description") || "").trim(),
                            announcementType: this._buildAnnouncementType(),
                            announcementStatus,
                            startAnnouncement,
                            endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt,
                            category: this._buildCategoryPayload()
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
                            ? (announcementStatus === "PUBLISHED" ? "duplicateSidebarSuccess" : "duplicateSidebarScheduledSuccess")
                            : (announcementStatus === "PUBLISHED" ? "createSidebarSuccess" : "createSidebarScheduledSuccess");
                        MessageToast.show(oBundle.getText(sKey, [oModel.getProperty("/title")]));
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
                        title: (oModel.getProperty("/title") || "").trim(),
                        description: (oModel.getProperty("/description") || "").trim(),
                        announcementType: this._buildAnnouncementType(),
                        announcementStatus,
                        startAnnouncement,
                        endAnnouncement,
                        publishedBy: sUserEmail,
                        publishedAt,
                        modifiedAt: sNow,
                        modifiedBy: sUserEmail,
                        category: this._buildCategoryPayload()
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
                            announcementStatus === "PUBLISHED" ? "updateSidebarSuccess" : "updateSidebarScheduledSuccess",
                            [oModel.getProperty("/title")]
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
                const sOrigDesc = oModel.getProperty("/originalDescription");
                const sOrigTitle = oModel.getProperty("/originalTitle");

                oModel.setProperty("/title", sOrigTitle);
                oModel.setProperty("/titleCharCount", `${sOrigTitle.length}/100`);
                oModel.setProperty("/categories", oModel.getProperty("/originalCategories").slice());
                oModel.setProperty("/description", sOrigDesc);
                oModel.setProperty("/descriptionCharCount", `${sOrigDesc.replace(/<[^>]*>/g, "").trim().length}/500`);
                oModel.setProperty("/popupAnnouncement", oModel.getProperty("/originalPopupAnnouncement"));
                oModel.setProperty("/publishDate", oModel.getProperty("/originalPublishDate"));
                oModel.setProperty("/expiryDate", oModel.getProperty("/originalExpiryDate"));
                oModel.setProperty("/publishLater", oModel.getProperty("/originalPublishLater"));
                this._oRichTextEditor?.setValue(sOrigDesc);
                MessageToast.show(oBundle.getText("resetToOriginalMessage"));

            } else if (oModel.getProperty("/isDuplicateMode")) {
                this._initSidebarModel();
                this._setDefaultDates(MODEL);
                oModel.setProperty("/isDuplicateMode", true);
                this._oRichTextEditor?.setValue("");
                MessageToast.show(oBundle.getText("formReset"));

            } else {
                this._initSidebarModel();
                this._setDefaultDates(MODEL);
                this._oRichTextEditor?.setValue("");
                MessageToast.show(oBundle.getText("formReset"));
            }

            oModel.setProperty("/resetButtonEnabled", false);
        }
    });
});