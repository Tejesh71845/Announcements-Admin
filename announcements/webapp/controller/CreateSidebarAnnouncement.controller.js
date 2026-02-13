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
                // Create mode - reset form
                this._initSidebarModel();
                setTimeout(() => {
                    this._initRichTextEditor("idRichTextCntr");
                }, 200);
            }
        },

        _loadAnnouncementForEdit: function (sAnnouncementId) {
            const oBusy = new sap.m.BusyDialog({ text: "Loading announcement..." });
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
                    MessageBox.error("Failed to load announcement data. Please try again.");
                    this._navBack();
                }
            });
        },

        _populateFormForEdit: function (oData) {
            const oSidebarModel = this.getView().getModel("sidebarModel");

            // Determine if popup based on announcement type
            const bIsPopup = oData.announcementType === "Sidebar (Popup)";

            // Extract category
            const aTypeIds = this._extractTypeIds(oData.toTypes);
            const sCategory = aTypeIds.length > 0 ? aTypeIds[0] : "";

            // Set model data
            oSidebarModel.setData({
                // Basic fields
                title: oData.title || "",
                category: sCategory,
                description: oData.description || "",
                popupAnnouncement: bIsPopup,

                // Publishing fields
                publishDate: oData.startAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "",
                expiryDate: oData.endAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "",
                publishLater: this._isPublishLater(oData.startAnnouncement),
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
                originalTitle: oData.title || "",
                originalCategory: sCategory,
                originalDescription: oData.description || "",
                originalPopupAnnouncement: bIsPopup,
                originalPublishDate: oData.startAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "",
                originalExpiryDate: oData.endAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "",
                originalPublishLater: this._isPublishLater(oData.startAnnouncement)
            });

            // Initialize RTE with description
            setTimeout(() => {
                this._initRichTextEditor("idRichTextCntr");
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
                category: "",
                description: "",
                popupAnnouncement: false,

                // Publishing fields
                publishDate: "",
                expiryDate: "",
                publishLater: false,
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
                            MessageToast.show(`Description cannot exceed ${MAX_CHARS} characters. Current: ${sPlainText.length}`);
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

            oModel.setProperty("/descriptionValueState", bValid ? "None" : "Error");
            oModel.setProperty("/descriptionValueStateText", bValid ? "" : "Description is required");

            const oContainer = this.byId("idRichTextCntr");
            if (oContainer) {
                if (!bValid) {
                    oContainer.addStyleClass("richTextErrorCss");
                } else {
                    oContainer.removeStyleClass("richTextErrorCss");
                }
            }
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

            this._handleResetButtonVisibility();
            this._validateField(oSource);
        },

        _validateField: function (oSource) {
            const oModel = this.getView().getModel("sidebarModel");
            const sId = oSource.getId();

            if (sId.indexOf("idTitleFld") > -1) {
                const sValue = (oModel.getProperty("/title") || "").trim();
                const bValid = sValue.length > 0;
                oModel.setProperty("/titleValueState", bValid ? "None" : "Error");
                oModel.setProperty("/titleValueStateText", bValid ? "" : "Title is required");
            }
        },

        onCategoryChange: function (oEvent) {
            const sSelectedKey = oEvent.getParameter("selectedItem").getKey();
            const oModel = this.getView().getModel("sidebarModel");

            oModel.setProperty("/category", sSelectedKey);
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

            oModel.setProperty("/publishDate", sValue);

            if (!sValue) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", "Publish Date is required");
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (oSelected < oToday) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", "Publish Date cannot be in the past");
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

            oModel.setProperty("/expiryDate", sValue);

            if (!sValue) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", "Expiry Date is required");
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (oSelected < oToday) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", "Expiry Date cannot be in the past");
                return;
            }

            const sPublishDate = oModel.getProperty("/publishDate");
            if (sPublishDate) {
                const oPublishDate = new Date(sPublishDate);
                oPublishDate.setHours(0, 0, 0, 0);
                if (oSelected <= oPublishDate) {
                    oModel.setProperty("/expiryDateValueState", "Error");
                    oModel.setProperty("/expiryDateValueStateText", "Expiry Date must be after Publish Date");
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
                const bChanged =
                    oModel.getProperty("/title") !== oModel.getProperty("/originalTitle") ||
                    oModel.getProperty("/description") !== oModel.getProperty("/originalDescription") ||
                    oModel.getProperty("/category") !== oModel.getProperty("/originalCategory") ||
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
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");

            let bValid = true;

            // Validate Title
            if (!sTitle) {
                oModel.setProperty("/titleValueState", "Error");
                oModel.setProperty("/titleValueStateText", "Title is required");
                bValid = false;
            } else {
                oModel.setProperty("/titleValueState", "None");
                oModel.setProperty("/titleValueStateText", "");
            }

            // Validate Description
            if (!sPlainText) {
                oModel.setProperty("/descriptionValueState", "Error");
                oModel.setProperty("/descriptionValueStateText", "Description is required");
                const oContainer = this.byId("idRichTextCntr");
                if (oContainer) {
                    oContainer.addStyleClass("richTextErrorCss");
                }
                bValid = false;
            } else {
                oModel.setProperty("/descriptionValueState", "None");
                oModel.setProperty("/descriptionValueStateText", "");
                const oContainer = this.byId("idRichTextCntr");
                if (oContainer) {
                    oContainer.removeStyleClass("richTextErrorCss");
                }
            }

            // Validate Publish Date
            if (!sPublishDate) {
                oModel.setProperty("/publishDateValueState", "Error");
                oModel.setProperty("/publishDateValueStateText", "Publish Date is required");
                bValid = false;
            }

            // Validate Expiry Date
            if (!sExpiryDate) {
                oModel.setProperty("/expiryDateValueState", "Error");
                oModel.setProperty("/expiryDateValueStateText", "Expiry Date is required");
                bValid = false;
            }

            return bValid;
        },

        onSubmitPress: function () {
            if (!this._validateAllFields()) {
                MessageToast.show("Please complete all required fields");
                return;
            }

            const oModel = this.getView().getModel("sidebarModel");
            const bIsEditMode = oModel.getProperty("/isEditMode");

            const sConfirmMsg = bIsEditMode
                ? "Are you sure you want to update this announcement?"
                : "Are you sure you want to submit this announcement?";

            MessageBox.confirm(sConfirmMsg, {
                title: bIsEditMode ? "Confirm Update" : "Confirm Submit",
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
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const sCategory = oModel.getProperty("/category");
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const bPopupAnnouncement = oModel.getProperty("/popupAnnouncement");
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

            // Determine announcement type based on popup toggle
            const sAnnouncementType = bPopupAnnouncement ? "Sidebar (Popup)" : "Sidebar";

            const oBusy = new sap.m.BusyDialog({ text: "Submitting announcement..." });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement;
                    const currentDateTime = new Date().toISOString();
                    const oToday = new Date();
                    oToday.setHours(0, 0, 0, 0);
                    const oPublishDate = new Date(sPublishDate);
                    oPublishDate.setHours(0, 0, 0, 0);

                    // Determine status based on publish later toggle and date
                    if (!bPublishLater && oPublishDate.getTime() === oToday.getTime()) {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sPublishDate).toISOString();
                    }

                    const oEndDate = new Date(sExpiryDate);
                    oEndDate.setDate(oEndDate.getDate() + 1);
                    endAnnouncement = oEndDate.toISOString();

                    const aToTypes = sCategory ? [{ type: { typeId: sCategory } }] : [];

                    const oPayload = {
                        data: [{
                            title: sTitle,
                            description: sDescriptionHTML,
                            announcementType: sAnnouncementType,
                            announcementStatus: announcementStatus,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt: currentDateTime,
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
                                        ? `Announcement '${sTitle}' published successfully!`
                                        : `Announcement '${sTitle}' scheduled for publication!`;
                                    MessageToast.show(sMessage);
                                    this._navBack();
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Create announcement failed:", status, err);
                                    let sErrorMessage = "Failed to create announcement. Please try again.";
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
                            MessageBox.error("Failed to initialize request. Please try again.");
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    MessageBox.error("Failed to get current user: " + error.message);
                });
        },

        _handleUpdate: function () {
            const oModel = this.getView().getModel("sidebarModel");
            const sEditId = oModel.getProperty("/editId");
            const sTitle = (oModel.getProperty("/title") || "").trim();
            const sCategory = oModel.getProperty("/category");
            const sDescriptionHTML = (oModel.getProperty("/description") || "").trim();
            const bPopupAnnouncement = oModel.getProperty("/popupAnnouncement");
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

            // Determine announcement type based on popup toggle
            const sAnnouncementType = bPopupAnnouncement ? "Sidebar (Popup)" : "Sidebar";

            const oBusy = new sap.m.BusyDialog({ text: "Updating announcement..." });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement;
                    const currentDateTime = new Date().toISOString();
                    const oToday = new Date();
                    oToday.setHours(0, 0, 0, 0);
                    const oPublishDate = new Date(sPublishDate);
                    oPublishDate.setHours(0, 0, 0, 0);

                    if (!bPublishLater && oPublishDate.getTime() === oToday.getTime()) {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sPublishDate).toISOString();
                    }

                    const oEndDate = new Date(sExpiryDate);
                    oEndDate.setDate(oEndDate.getDate() + 1);
                    endAnnouncement = oEndDate.toISOString();

                    const aToTypes = sCategory ? [{ type: { typeId: sCategory } }] : [];

                    const oPayload = {
                        title: sTitle,
                        description: sDescriptionHTML,
                        announcementType: sAnnouncementType,
                        announcementStatus: announcementStatus,
                        startAnnouncement: startAnnouncement,
                        endAnnouncement: endAnnouncement,
                        publishedBy: sUserEmail,
                        publishedAt: currentDateTime,
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
                                        ? `Announcement '${sTitle}' updated and published successfully!`
                                        : `Announcement '${sTitle}' updated and scheduled for publication!`;
                                    MessageToast.show(sMessage);
                                    this._navBack();
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Update announcement failed:", status, err);
                                    let sErrorMessage = "Failed to update announcement. Please try again.";
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
                            MessageBox.error("Failed to initialize request. Please try again.");
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    MessageBox.error("Failed to get current user: " + error.message);
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
            MessageBox.confirm(
                "Are you sure you want to reset the form? All entered data will be lost.",
                {
                    title: "Confirm Reset",
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
            const bIsEditMode = oModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Reset to original values
                oModel.setProperty("/title", oModel.getProperty("/originalTitle"));
                oModel.setProperty("/category", oModel.getProperty("/originalCategory"));
                oModel.setProperty("/description", oModel.getProperty("/originalDescription"));
                oModel.setProperty("/popupAnnouncement", oModel.getProperty("/originalPopupAnnouncement"));
                oModel.setProperty("/publishDate", oModel.getProperty("/originalPublishDate"));
                oModel.setProperty("/expiryDate", oModel.getProperty("/originalExpiryDate"));
                oModel.setProperty("/publishLater", oModel.getProperty("/originalPublishLater"));

                // Update RTE
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue(oModel.getProperty("/originalDescription"));
                }

                MessageToast.show("Reset to original values");
            } else {
                // Clear all fields
                this._initSidebarModel();
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue("");
                }
                MessageToast.show("Form has been reset");
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