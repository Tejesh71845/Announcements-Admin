sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/richtexteditor/RichTextEditor",
    "sap/ui/richtexteditor/library",
    "com/incture/announcements/utils/formatter"
], (Controller, JSONModel, MessageToast, MessageBox, RTE, library, formatter) => {
    "use strict";

    return Controller.extend("com.incture.announcements.controller.CreateSidebarAnnouncement", {

        formatter: formatter,

        onInit: function () {
            this._router = this.getOwnerComponent().getRouter();
            this._router.getRoute("CreateSidebarAnnouncement").attachPatternMatched(this._onRouteMatched, this);

            // Initialize models
            this._initSidebarModel();
            this._initCategoryModel();
        },

        _onRouteMatched: function (oEvent) {
            const oArgs = oEvent.getParameter("arguments");
            const sEditId = oArgs ? oArgs.announcementId : null;

            if (sEditId) {
                // Edit mode - load announcement data
                this._loadAnnouncementForEdit(sEditId);
            } else {
                // Create mode - reset form and set default dates
                this._initSidebarModel();
                this._setDefaultDates();
                setTimeout(() => {
                    this._initRichTextEditor("idSidebarRichTextCntr");
                }, 200);
            }
        },

        _setDefaultDates: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            // Set publish date to today
            const sTodayValue = formatter.formatDateToValue(oToday);
            oModel.setProperty("/publishDate", sTodayValue);

            // Set expiry date to today + 30 days
            const oExpiryDate = new Date(oToday);
            oExpiryDate.setDate(oExpiryDate.getDate() + 30);
            const sExpiryValue = formatter.formatDateToValue(oExpiryDate);
            oModel.setProperty("/expiryDate", sExpiryValue);

            // Update min expiry date
            const oMinExpiry = new Date(oToday);
            oMinExpiry.setDate(oMinExpiry.getDate() + 1);
            oModel.setProperty("/minExpiryDate", oMinExpiry);

            // Publish date is disabled by default (publish today)
            oModel.setProperty("/publishDateEnabled", false);
            oModel.setProperty("/showPublishTodayText", true);
        },

        _loadAnnouncementForEdit: function (sAnnouncementId) {
            const oBusy = new sap.m.BusyDialog({
                text: this.getView().getModel("i18n").getResourceBundle().getText("loadingAnnouncement")
            });
            oBusy.open();

            const oModel = this.getOwnerComponent().getModel("announcementModel");

            oModel.read(`/Announcements('${sAnnouncementId}')`, {
                urlParameters: {
                    "$expand": "toTypes"
                },
                success: (oData) => {
                    oBusy.close();
                    this._populateFormForEdit(oData);
                },
                error: (oError) => {
                    oBusy.close();
                    console.error("Failed to load announcement:", oError);
                    const sErrorMsg = this.getView().getModel("i18n").getResourceBundle().getText("loadAnnouncementError");
                    MessageBox.error(sErrorMsg);
                    this._navBack();
                }
            });
        },

        _populateFormForEdit: function (oData) {
            const oSidebarModel = this.getView().getModel("sidebarModel");

            // Determine if popup based on announcement type
            const bIsPopup = oData.announcementType === "Sidebar (Popup)";

            // Extract categories (now multiple)
            const aTypeIds = this._extractTypeIds(oData.toTypes);

            // Parse dates
            const sPublishDate = oData.startAnnouncement ?
                formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "";
            const sExpiryDate = oData.endAnnouncement ?
                formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "";

            const bPublishLater = this._isPublishLater(oData.startAnnouncement);

            // Calculate character counts
            const sTitle = oData.title || "";
            const sDescription = oData.description || "";
            const sDescPlainText = sDescription.replace(/<[^>]*>/g, "").trim();

            // Set model data
            oSidebarModel.setData({
                // Basic fields
                title: sTitle,
                titleCharCount: `${sTitle.length}/100`,
                categories: aTypeIds,
                description: sDescription,
                descriptionCharCount: `${sDescPlainText.length}/500`,
                popupAnnouncement: bIsPopup,

                // Publishing fields
                publishDate: sPublishDate,
                expiryDate: sExpiryDate,
                publishLater: bPublishLater,
                publishDateEnabled: bPublishLater,
                showPublishTodayText: !bPublishLater,
                minPublishDate: new Date(),
                minExpiryDate: new Date(),

                // Audience metadata (static UI fields)
                organizationFinance: true,
                organizationNonFinance: true,
                sectorCorporate: true,
                sectorInnovativeMedicine: true,
                sectorMedTech: true,

                // Validation states
                titleValueState: "None",
                titleValueStateText: "",
                descriptionValueState: "None",
                descriptionValueStateText: "",
                publishDateValueState: "None",
                publishDateValueStateText: "",
                expiryDateValueState: "None",
                expiryDateValueStateText: "",

                // Button visibility
                showResetButton: false,

                // Edit mode
                isEditMode: true,
                editId: oData.announcementId,

                // Store original values for reset
                originalTitle: sTitle,
                originalCategories: aTypeIds.slice(), // Clone array
                originalDescription: sDescription,
                originalPopupAnnouncement: bIsPopup,
                originalPublishDate: sPublishDate,
                originalExpiryDate: sExpiryDate,
                originalPublishLater: bPublishLater
            });

            // Initialize RTE with description
            setTimeout(() => {
                this._initRichTextEditor("idSidebarRichTextCntr");
            }, 200);
        },

        _extractTypeIds: function (toTypes) {
            if (!toTypes) {
                return [];
            }

            let aResults = toTypes.__list || toTypes.results || toTypes;

            if (!Array.isArray(aResults)) {
                return [];
            }

            return aResults
                .map(item => {
                    if (item.type_typeId) {
                        return item.type_typeId;
                    }
                    if (item.type && item.type.typeId) {
                        return item.type.typeId;
                    }
                    if (item.typeId) {
                        return item.typeId;
                    }
                    return null;
                })
                .filter(Boolean);
        },

        _isPublishLater: function (sStartAnnouncement) {
            if (!sStartAnnouncement) return false;

            const oStartDate = new Date(sStartAnnouncement);
            oStartDate.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            return oStartDate > oToday;
        },

        _initSidebarModel: function () {
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            const oModel = new JSONModel({
                // Basic fields
                title: "",
                titleCharCount: "0/100",
                categories: [],
                description: "",
                descriptionCharCount: "0/500",
                popupAnnouncement: false,

                // Publishing fields
                publishDate: "",
                expiryDate: "",
                publishLater: false,
                publishDateEnabled: false,
                showPublishTodayText: true,
                minPublishDate: oToday,
                minExpiryDate: oToday,

                // Audience metadata (static UI fields)
                organizationFinance: true,
                organizationNonFinance: true,
                sectorCorporate: true,
                sectorInnovativeMedicine: true,
                sectorMedTech: true,

                // Validation states
                titleValueState: "None",
                titleValueStateText: "",
                descriptionValueState: "None",
                descriptionValueStateText: "",
                publishDateValueState: "None",
                publishDateValueStateText: "",
                expiryDateValueState: "None",
                expiryDateValueStateText: "",

                // Button visibility
                showResetButton: false,

                // Edit mode
                isEditMode: false,
                editId: null
            });

            this.getView().setModel(oModel, "sidebarModel");
        },

        _initCategoryModel: function () {
            const oCategoryModel = new JSONModel({
                category: [],
                idToNameMap: {},
                nameToIdMap: {}
            });
            this.getView().setModel(oCategoryModel, "categoryModel");

            const oDataModel = this.getOwnerComponent().getModel("typeModel");
            this._loadCategories(oDataModel, oCategoryModel);
        },

        _loadCategories: function (oDataModel, oCategoryModel) {
            this.getView().setBusy(true);

            oDataModel.read("/Types", {
                urlParameters: {
                    "$select": "typeId,name,description"
                },
                success: (oData) => {
                    this.getView().setBusy(false);

                    const aEntries = oData.results || [];
                    const aDropdownData = [];
                    const idToName = {};
                    const nameToId = {};

                    aEntries.forEach((entry) => {
                        const typeId = entry.typeId;
                        const typeName = entry.name;

                        if (typeId && typeName) {
                            aDropdownData.push({
                                key: typeId,
                                text: typeName
                            });
                            idToName[typeId] = typeName;
                            nameToId[typeName] = typeId;
                        }
                    });

                    oCategoryModel.setProperty("/category", aDropdownData);
                    oCategoryModel.setProperty("/idToNameMap", idToName);
                    oCategoryModel.setProperty("/nameToIdMap", nameToId);
                },
                error: (oError) => {
                    this.getView().setBusy(false);
                    console.error("Failed to fetch categories:", oError);
                    MessageBox.error("Failed to load categories");
                }
            });
        },

        _initRichTextEditor: function (sContainerId) {
            if (this._oRichTextEditor) {
                this._oRichTextEditor.destroy();
            }

            const EditorType = library.EditorType;
            const oModel = this.getView().getModel("sidebarModel");
            const sDescription = oModel.getProperty("/description") || "";

            this._oRichTextEditor = new RTE({
                editorType: EditorType.TinyMCE7,
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
                            const oBundle = this.getView().getModel("i18n").getResourceBundle();
                            const sMsg = oBundle.getText("descriptionCharLimitError", [MAX_CHARS, sPlainText.length]);
                            MessageToast.show(sMsg);
                            this._isCharLimitToastShown = true;
                            setTimeout(() => {
                                this._isCharLimitToastShown = false;
                            }, 3000);
                        }

                        const sPreviousValue = oModel.getProperty("/description") || "";
                        this._oRichTextEditor.setValue(sPreviousValue);
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
            const oModel = this.getView().getModel("sidebarModel");
            const sDescriptionHTML = oModel.getProperty("/description") || "";
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();
            const bValid = sPlainText.length > 0;

            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            oModel.setProperty("/descriptionValueState", bValid ? "None" : "Error");
            oModel.setProperty("/descriptionValueStateText", bValid ? "" : oBundle.getText("descriptionRequired"));

            const oContainer = this.byId("idSidebarRichTextCntr");
            if (oContainer) {
                if (!bValid) {
                    oContainer.addStyleClass("richTextErrorCss");
                } else {
                    oContainer.removeStyleClass("richTextErrorCss");
                }
            }
        },

        onTitleLiveChange: function (oEvent) {
            const sValue = oEvent.getParameter("value") || "";
            const oModel = this.getView().getModel("sidebarModel");

            // Update character count
            oModel.setProperty("/titleCharCount", `${sValue.length}/100`);
        },

        onInputChange: function (oEvent) {
            const oSource = oEvent.getSource();
            let sValue = oEvent.getParameter("value") || "";

            // Remove special characters
            sValue = sValue.replace(/[^a-zA-Z0-9 ]/g, "");

            const iMaxLength = oSource.getMaxLength && oSource.getMaxLength() > 0
                ? oSource.getMaxLength()
                : 100;

            if (sValue.length > iMaxLength) {
                sValue = sValue.substring(0, iMaxLength);
                oSource.setValue(sValue);
            }

            if (sValue !== oEvent.getParameter("value")) {
                oSource.setValue(sValue);
            }

            // Update character count
            const oModel = this.getView().getModel("sidebarModel");
            oModel.setProperty("/titleCharCount", `${sValue.length}/100`);

            this._handleResetButtonVisibility();
            this._validateField(oSource);
        },

        _validateField: function (oSource) {
            const oModel = this.getView().getModel("sidebarModel");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sId = oSource.getId();

            if (sId.indexOf("idSidebarTitleFld") > -1) {
                const sValue = (oModel.getProperty("/title") || "").trim();
                const bValid = sValue.length > 0;
                oModel.setProperty("/titleValueState", bValid ? "None" : "Error");
                oModel.setProperty("/titleValueStateText", bValid ? "" : oBundle.getText("titleRequired"));
            }
        },

        onCategoryChange: function (oEvent) {
            const oSource = oEvent.getSource();
            const aSelectedKeys = oSource.getSelectedKeys();
            const oModel = this.getView().getModel("sidebarModel");

            oModel.setProperty("/categories", aSelectedKeys);
            this._handleResetButtonVisibility();
        },

        onPopupSwitchChange: function (oEvent) {
            const bState = oEvent.getParameter("state");
            const oModel = this.getView().getModel("sidebarModel");

            oModel.setProperty("/popupAnnouncement", bState);
            this._handleResetButtonVisibility();
        },

        onPublishDateChange: function (oEvent) {
            const sValue = oEvent.getParameter("value");
            const oModel = this.getView().getModel("sidebarModel");
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

            // Auto-set expiry date to 30 days later
            const oExpiryDate = new Date(oSelected);
            oExpiryDate.setDate(oExpiryDate.getDate() + 30);
            const sExpiryValue = formatter.formatDateToValue(oExpiryDate);

            oModel.setProperty("/expiryDate", sExpiryValue);
            oModel.setProperty("/expiryDateValueState", "None");
            oModel.setProperty("/expiryDateValueStateText", "");

            // Update min expiry date
            const oMinExpiry = new Date(oSelected);
            oMinExpiry.setDate(oMinExpiry.getDate() + 1);
            oModel.setProperty("/minExpiryDate", oMinExpiry);

            this._handleResetButtonVisibility();
        },

        onExpiryDateChange: function (oEvent) {
            const sValue = oEvent.getParameter("value");
            const oModel = this.getView().getModel("sidebarModel");
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
                const oPublishDate = new Date(sPublishDate);
                oPublishDate.setHours(0, 0, 0, 0);
                if (oSelected <= oPublishDate) {
                    oModel.setProperty("/expiryDateValueState", "Error");
                    oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDateBeforePublishError"));
                    return;
                }
            }

            oModel.setProperty("/expiryDateValueState", "None");
            oModel.setProperty("/expiryDateValueStateText", "");

            this._handleResetButtonVisibility();
        },

        onPublishLaterChange: function (oEvent) {
            const bState = oEvent.getParameter("state");
            const oModel = this.getView().getModel("sidebarModel");

            oModel.setProperty("/publishLater", bState);

            if (bState) {
                // Enable publish date and set to tomorrow
                const oTomorrow = new Date();
                oTomorrow.setDate(oTomorrow.getDate() + 1);
                oTomorrow.setHours(0, 0, 0, 0);

                const sTomorrowValue = formatter.formatDateToValue(oTomorrow);
                oModel.setProperty("/publishDate", sTomorrowValue);
                oModel.setProperty("/publishDateEnabled", true);
                oModel.setProperty("/showPublishTodayText", false);

                // Update expiry date to tomorrow + 30 days
                const oExpiryDate = new Date(oTomorrow);
                oExpiryDate.setDate(oExpiryDate.getDate() + 30);
                const sExpiryValue = formatter.formatDateToValue(oExpiryDate);
                oModel.setProperty("/expiryDate", sExpiryValue);

                // Update min expiry date
                const oMinExpiry = new Date(oTomorrow);
                oMinExpiry.setDate(oMinExpiry.getDate() + 1);
                oModel.setProperty("/minExpiryDate", oMinExpiry);
            } else {
                // Disable publish date and set to today
                const oToday = new Date();
                oToday.setHours(0, 0, 0, 0);

                const sTodayValue = formatter.formatDateToValue(oToday);
                oModel.setProperty("/publishDate", sTodayValue);
                oModel.setProperty("/publishDateEnabled", false);
                oModel.setProperty("/showPublishTodayText", true);

                // Update expiry date to today + 30 days
                const oExpiryDate = new Date(oToday);
                oExpiryDate.setDate(oExpiryDate.getDate() + 30);
                const sExpiryValue = formatter.formatDateToValue(oExpiryDate);
                oModel.setProperty("/expiryDate", sExpiryValue);

                // Update min expiry date
                const oMinExpiry = new Date(oToday);
                oMinExpiry.setDate(oMinExpiry.getDate() + 1);
                oModel.setProperty("/minExpiryDate", oMinExpiry);
            }

            this._handleResetButtonVisibility();
        },

        onOrganizationChange: function () {
            this._handleResetButtonVisibility();
        },

        onSectorChange: function () {
            this._handleResetButtonVisibility();
        },

        _handleResetButtonVisibility: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const bIsEditMode = oModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Edit mode - check if any field changed from original
                const aCategories = oModel.getProperty("/categories") || [];
                const aOriginalCategories = oModel.getProperty("/originalCategories") || [];

                const bCategoriesChanged = aCategories.length !== aOriginalCategories.length ||
                    !aCategories.every(cat => aOriginalCategories.includes(cat));

                const bChanged =
                    oModel.getProperty("/title") !== oModel.getProperty("/originalTitle") ||
                    oModel.getProperty("/description") !== oModel.getProperty("/originalDescription") ||
                    bCategoriesChanged ||
                    oModel.getProperty("/popupAnnouncement") !== oModel.getProperty("/originalPopupAnnouncement") ||
                    oModel.getProperty("/publishDate") !== oModel.getProperty("/originalPublishDate") ||
                    oModel.getProperty("/expiryDate") !== oModel.getProperty("/originalExpiryDate") ||
                    oModel.getProperty("/publishLater") !== oModel.getProperty("/originalPublishLater");

                oModel.setProperty("/showResetButton", bChanged);
            } else {
                // Create mode - check if any data entered
                const sTitle = oModel.getProperty("/title") || "";
                const sDescription = oModel.getProperty("/description") || "";
                const sPublishDate = oModel.getProperty("/publishDate") || "";
                const sExpiryDate = oModel.getProperty("/expiryDate") || "";

                const bHasData = sTitle.length > 0 ||
                    sDescription.length > 0 ||
                    sPublishDate.length > 0 ||
                    sExpiryDate.length > 0;

                oModel.setProperty("/showResetButton", bHasData);
            }
        },

        _validateAllFields: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");

            let bValid = true;

            // Validate Title
            if (!sTitle) {
                oModel.setProperty("/titleValueState", "Error");
                oModel.setProperty("/titleValueStateText", oBundle.getText("titleRequired"));
                bValid = false;
            } else {
                oModel.setProperty("/titleValueState", "None");
                oModel.setProperty("/titleValueStateText", "");
            }

            // Validate Description
            if (!sPlainText) {
                oModel.setProperty("/descriptionValueState", "Error");
                oModel.setProperty("/descriptionValueStateText", oBundle.getText("descriptionRequired"));
                const oContainer = this.byId("idSidebarRichTextCntr");
                if (oContainer) {
                    oContainer.addStyleClass("richTextErrorCss");
                }
                bValid = false;
            } else {
                oModel.setProperty("/descriptionValueState", "None");
                oModel.setProperty("/descriptionValueStateText", "");
                const oContainer = this.byId("idSidebarRichTextCntr");
                if (oContainer) {
                    oContainer.removeStyleClass("richTextErrorCss");
                }
            }

            // Validate Publish Date
            if (!sPublishDate) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", oBundle.getText("publishDateRequired"));
                bValid = false;
            }

            // Validate Expiry Date
            if (!sExpiryDate) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", oBundle.getText("expiryDateRequired"));
                bValid = false;
            }

            return bValid;
        },

        onSubmitPress: function () {
            if (!this._validateAllFields()) {
                const oBundle = this.getView().getModel("i18n").getResourceBundle();
                MessageToast.show(oBundle.getText("validationError"));
                return;
            }

            const oModel = this.getView().getModel("sidebarModel");
            const bIsEditMode = oModel.getProperty("/isEditMode");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();

            const sConfirmMsg = bIsEditMode
                ? oBundle.getText("updateConfirmMessage")
                : oBundle.getText("submitConfirmMessage");

            const sTitle = bIsEditMode
                ? oBundle.getText("updateConfirmTitle")
                : oBundle.getText("submitConfirmTitle");

            MessageBox.confirm(sConfirmMsg, {
                title: sTitle,
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.YES,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        if (bIsEditMode) {
                            this._handleUpdate();
                        } else {
                            this._handleSubmit();
                        }
                    }
                }
            });
        },

        _handleSubmit: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const aCategories = oModel.getProperty("/categories") || [];
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const bPopupAnnouncement = oModel.getProperty("/popupAnnouncement");
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

            // Determine announcement type based on popup toggle
            const sAnnouncementType = bPopupAnnouncement ? "Sidebar (Popup)" : "Sidebar";

            const oBusy = new sap.m.BusyDialog({
                text: oBundle.getText("submittingAnnouncement")
            });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement, publishedAt;
                    const oToday = new Date();
                    oToday.setHours(0, 0, 0, 0);
                    const oPublishDate = new Date(sPublishDate);
                    oPublishDate.setHours(0, 0, 0, 0);

                    // Determine status and dates based on publish later toggle
                    if (!bPublishLater && oPublishDate.getTime() === oToday.getTime()) {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = new Date().toISOString();
                        publishedAt = new Date().toISOString();
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sPublishDate).toISOString();
                        publishedAt = new Date(sPublishDate).toISOString();
                    }

                    const oEndDate = new Date(sExpiryDate);
                    oEndDate.setDate(oEndDate.getDate() + 1);
                    endAnnouncement = oEndDate.toISOString();

                    // Build toTypes array from selected categories
                    const aToTypes = aCategories.map(catId => ({ type: { typeId: catId } }));

                    const oPayload = {
                        data: [{
                            title: sTitle,
                            description: sDescriptionHTML,
                            announcementType: sAnnouncementType,
                            announcementStatus: announcementStatus,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt: publishedAt,
                            toTypes: aToTypes
                        }]
                    };

                    this._getCSRFToken()
                        .then((csrfToken) => {
                            $.ajax({
                                url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/bulkCreateAnnouncements",
                                method: "POST",
                                contentType: "application/json",
                                dataType: "json",
                                headers: {
                                    "X-CSRF-Token": csrfToken
                                },
                                data: JSON.stringify(oPayload),
                                success: (oResponse) => {
                                    oBusy.close();
                                    const sMessage = announcementStatus === "PUBLISHED"
                                        ? oBundle.getText("createSidebarSuccess", [sTitle])
                                        : oBundle.getText("createSidebarScheduledSuccess", [sTitle]);
                                    MessageToast.show(sMessage);
                                    this._navBack();
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Create announcement failed:", status, err);
                                    let sErrorMessage = oBundle.getText("createAnnouncementError");
                                    if (xhr.responseJSON?.error?.message) {
                                        sErrorMessage = xhr.responseJSON.error.message;
                                    }
                                    MessageBox.error(sErrorMessage);
                                }
                            });
                        })
                        .catch((err) => {
                            oBusy.close();
                            console.error("CSRF token fetch failed:", err);
                            MessageBox.error(oBundle.getText("csrfTokenError"));
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    MessageBox.error(oBundle.getText("getCurrentUserError", [error.message]));
                });
        },

        _handleUpdate: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const sEditId = oModel.getProperty("/editId");
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const aCategories = oModel.getProperty("/categories") || [];
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const bPopupAnnouncement = oModel.getProperty("/popupAnnouncement");
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

            // Determine announcement type based on popup toggle
            const sAnnouncementType = bPopupAnnouncement ? "Sidebar (Popup)" : "Sidebar";

            const oBusy = new sap.m.BusyDialog({
                text: oBundle.getText("updatingAnnouncement")
            });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement, publishedAt;
                    const currentDateTime = new Date().toISOString();
                    const oToday = new Date();
                    oToday.setHours(0, 0, 0, 0);
                    const oPublishDate = new Date(sPublishDate);
                    oPublishDate.setHours(0, 0, 0, 0);

                    if (!bPublishLater && oPublishDate.getTime() === oToday.getTime()) {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                        publishedAt = currentDateTime;
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sPublishDate).toISOString();
                        publishedAt = new Date(sPublishDate).toISOString();
                    }

                    const oEndDate = new Date(sExpiryDate);
                    oEndDate.setDate(oEndDate.getDate() + 1);
                    endAnnouncement = oEndDate.toISOString();

                    // Build toTypes array from selected categories
                    const aToTypes = aCategories.map(catId => ({ type: { typeId: catId } }));

                    const oPayload = {
                        title: sTitle,
                        description: sDescriptionHTML,
                        announcementType: sAnnouncementType,
                        announcementStatus: announcementStatus,
                        startAnnouncement: startAnnouncement,
                        endAnnouncement: endAnnouncement,
                        publishedBy: sUserEmail,
                        publishedAt: publishedAt,
                        modifiedAt: currentDateTime,
                        modifiedBy: sUserEmail,
                        toTypes: aToTypes
                    };

                    this._getCSRFToken()
                        .then((csrfToken) => {
                            $.ajax({
                                url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${sEditId}')`,
                                method: "PATCH",
                                contentType: "application/json",
                                dataType: "json",
                                headers: {
                                    "X-CSRF-Token": csrfToken
                                },
                                data: JSON.stringify(oPayload),
                                success: (oResponse) => {
                                    oBusy.close();
                                    const sMessage = announcementStatus === "PUBLISHED"
                                        ? oBundle.getText("updateSidebarSuccess", [sTitle])
                                        : oBundle.getText("updateSidebarScheduledSuccess", [sTitle]);
                                    MessageToast.show(sMessage);
                                    this._navBack();
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Update announcement failed:", status, err);
                                    let sErrorMessage = oBundle.getText("updateAnnouncementError");
                                    if (xhr.responseJSON?.error?.message) {
                                        sErrorMessage = xhr.responseJSON.error.message;
                                    }
                                    MessageBox.error(sErrorMessage);
                                }
                            });
                        })
                        .catch((err) => {
                            oBusy.close();
                            console.error("CSRF token fetch failed:", err);
                            MessageBox.error(oBundle.getText("csrfTokenError"));
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    MessageBox.error(oBundle.getText("getCurrentUserError", [error.message]));
                });
        },

        _getCSRFToken: function () {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/",
                    method: "GET",
                    headers: {
                        "X-CSRF-Token": "Fetch"
                    },
                    success: function (data, textStatus, request) {
                        const token = request.getResponseHeader("X-CSRF-Token");
                        resolve(token);
                    },
                    error: function (xhr, status, err) {
                        console.error("CSRF token fetch failed:", status, err);
                        reject(err);
                    }
                });
            });
        },

        getCurrentUserEmail: async function () {
            try {
                const appId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
                const appPath = appId.replaceAll(".", "/");
                const appModulePath = jQuery.sap.getModulePath(appPath);
                const url = appModulePath + "/user-api/currentUser";
                const oModel = new JSONModel();
                await oModel.loadData(url);

                const data = oModel.getData();
                if (data && data.email) {
                    return data.email;
                } else {
                    throw new Error("Email not found in the response.");
                }
            } catch (error) {
                throw new Error("Failed to fetch current user: " + error.message);
            }
        },

        onResetPress: function () {
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            MessageBox.confirm(
                oBundle.getText("resetConfirmMessage"),
                {
                    title: oBundle.getText("resetConfirmTitle"),
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.NO,
                    onClose: (oAction) => {
                        if (oAction === MessageBox.Action.YES) {
                            this._performReset();
                        }
                    }
                }
            );
        },

        _performReset: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const oBundle = this.getView().getModel("i18n").getResourceBundle();
            const bIsEditMode = oModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Reset to original values
                const sOriginalTitle = oModel.getProperty("/originalTitle");
                const sOriginalDescription = oModel.getProperty("/originalDescription");
                const sOriginalDescPlainText = sOriginalDescription.replace(/<[^>]*>/g, "").trim();

                oModel.setProperty("/title", sOriginalTitle);
                oModel.setProperty("/titleCharCount", `${sOriginalTitle.length}/100`);
                oModel.setProperty("/categories", oModel.getProperty("/originalCategories").slice());
                oModel.setProperty("/description", sOriginalDescription);
                oModel.setProperty("/descriptionCharCount", `${sOriginalDescPlainText.length}/500`);
                oModel.setProperty("/popupAnnouncement", oModel.getProperty("/originalPopupAnnouncement"));
                oModel.setProperty("/publishDate", oModel.getProperty("/originalPublishDate"));
                oModel.setProperty("/expiryDate", oModel.getProperty("/originalExpiryDate"));
                oModel.setProperty("/publishLater", oModel.getProperty("/originalPublishLater"));

                // Update RTE
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue(sOriginalDescription);
                }

                MessageToast.show(oBundle.getText("resetToOriginalMessage"));
            } else {
                // Clear all fields
                this._initSidebarModel();
                this._setDefaultDates();
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue("");
                }
                MessageToast.show(oBundle.getText("formReset"));
            }

            oModel.setProperty("/showResetButton", false);
        },

        onNavBack: function () {
            this._navBack();
        },

        _navBack: function () {
            this._router.navTo("Announcement");
        }
    });
});