sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "com/incture/announcements/utils/formatter"
], (Controller, JSONModel, MessageToast, MessageBox, formatter) => {
    "use strict";

    return Controller.extend("com.incture.announcements.controller.CreateBannerAnnouncement", {

        formatter: formatter,

        onInit: function () {
            this._router = this.getOwnerComponent().getRouter();
            this._router.getRoute("CreateBannerAnnouncement").attachPatternMatched(this._onRouteMatched, this);

            this._initBannerModel();
        },

        _onRouteMatched: function (oEvent) {
            const oArgs = oEvent.getParameter("arguments");
            const sEditId = oArgs ? oArgs.announcementId : null;

            if (sEditId) {
                // Edit mode - load announcement data
                this._loadAnnouncementForEdit(sEditId);
            } else {
                // Create mode - reset form
                this._initBannerModel();
            }
        },

        _loadAnnouncementForEdit: function (sAnnouncementId) {
            const oBusy = new sap.m.BusyDialog({ text: "Loading announcement..." });
            oBusy.open();

            const oModel = this.getOwnerComponent().getModel("announcementModel");

            oModel.read(`/Announcements('${sAnnouncementId}')`, {
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
            const oBannerModel = this.getView().getModel("bannerModel");

            oBannerModel.setData({
                // Basic fields - For Banner, title is the content
                announcementContent: oData.title || "",

                // Publishing fields
                publishDate: oData.startAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "",
                expiryDate: oData.endAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "",
                publishLater: this._isPublishLater(oData.startAnnouncement),
                minPublishDate: new Date(),
                minExpiryDate: new Date(),

                // Validation states
                announcementContentValueState: "None",
                announcementContentValueStateText: "",
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
                originalAnnouncementContent: oData.title || "",
                originalPublishDate: oData.startAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.startAnnouncement)) : "",
                originalExpiryDate: oData.endAnnouncement ?
                    formatter.formatDateToValue(new Date(oData.endAnnouncement)) : "",
                originalPublishLater: this._isPublishLater(oData.startAnnouncement)
            });
        },

        _isPublishLater: function (sStartAnnouncement) {
            if (!sStartAnnouncement) return false;

            const oStartDate = new Date(sStartAnnouncement);
            oStartDate.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            return oStartDate > oToday;
        },

        _initBannerModel: function () {
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            const oModel = new JSONModel({
                // Basic fields
                announcementContent: "",

                // Publishing fields
                publishDate: "",
                expiryDate: "",
                publishLater: false,
                minPublishDate: oToday,
                minExpiryDate: oToday,

                // Validation states
                announcementContentValueState: "None",
                announcementContentValueStateText: "",
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

            this.getView().setModel(oModel, "bannerModel");
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
            const oModel = this.getView().getModel("bannerModel");
            const sId = oSource.getId();

            if (sId.indexOf("idAnnouncementContentFld") > -1) {
                const sValue = (oModel.getProperty("/announcementContent") || "").trim();
                const bValid = sValue.length > 0;
                oModel.setProperty("/announcementContentValueState", bValid ? "None" : "Error");
                oModel.setProperty("/announcementContentValueStateText", bValid ? "" : "Announcement Content is required");
            }
        },

        onPublishDateChange: function (oEvent) {
            const sValue = oEvent.getParameter("value");
            const oModel = this.getView().getModel("bannerModel");

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
            const oModel = this.getView().getModel("bannerModel");

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
            const oModel = this.getView().getModel("bannerModel");

            oModel.setProperty("/publishLater", bState);
            this._handleResetButtonVisibility();
        },

        _handleResetButtonVisibility: function () {
            const oModel = this.getView().getModel("bannerModel");
            const bIsEditMode = oModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Edit mode - check if any field changed from original
                const bChanged =
                    oModel.getProperty("/announcementContent") !== oModel.getProperty("/originalAnnouncementContent") ||
                    oModel.getProperty("/publishDate") !== oModel.getProperty("/originalPublishDate") ||
                    oModel.getProperty("/expiryDate") !== oModel.getProperty("/originalExpiryDate") ||
                    oModel.getProperty("/publishLater") !== oModel.getProperty("/originalPublishLater");

                oModel.setProperty("/showResetButton", bChanged);
            } else {
                // Create mode - check if any data entered
                const sContent = oModel.getProperty("/announcementContent") || "";
                const sPublishDate = oModel.getProperty("/publishDate") || "";
                const sExpiryDate = oModel.getProperty("/expiryDate") || "";

                const bHasData = sContent.length > 0 ||
                    sPublishDate.length > 0 ||
                    sExpiryDate.length > 0;

                oModel.setProperty("/showResetButton", bHasData);
            }
        },

        _validateAllFields: function () {
            const oModel = this.getView().getModel("bannerModel");
            const sContent = (oModel.getProperty("/announcementContent") || "").trim();
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");

            let bValid = true;

            // Validate Content
            if (!sContent) {
                oModel.setProperty("/announcementContentValueState", "Error");
                oModel.setProperty("/announcementContentValueStateText", "Announcement Content is required");
                bValid = false;
            } else {
                oModel.setProperty("/announcementContentValueState", "None");
                oModel.setProperty("/announcementContentValueStateText", "");
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

            const oModel = this.getView().getModel("bannerModel");
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
            const oModel = this.getView().getModel("bannerModel");
            const sContent = (oModel.getProperty("/announcementContent") || "").trim();
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

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

                    const oPayload = {
                        data: [{
                            title: sContent,  // Announcement Content goes in title
                            description: "",  // Empty description for banner
                            announcementType: "Banner",
                            announcementStatus: announcementStatus,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt: currentDateTime,
                            toTypes: []
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
                                        ? "Banner announcement published successfully!"
                                        : "Banner announcement scheduled for publication!";
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
            const oModel = this.getView().getModel("bannerModel");
            const sEditId = oModel.getProperty("/editId");
            const sContent = (oModel.getProperty("/announcementContent") || "").trim();
            const sPublishDate = oModel.getProperty("/publishDate");
            const sExpiryDate = oModel.getProperty("/expiryDate");
            const bPublishLater = oModel.getProperty("/publishLater");

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

                    const oPayload = {
                        title: sContent,  // Announcement Content goes in title
                        description: "Banner Announcement",  // Empty description for banner
                        announcementType: "Banner",
                        announcementStatus: announcementStatus,
                        startAnnouncement: startAnnouncement,
                        endAnnouncement: endAnnouncement,
                        publishedBy: sUserEmail,
                        publishedAt: currentDateTime,
                        modifiedAt: currentDateTime,
                        modifiedBy: sUserEmail,
                        toTypes: []
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
                                        ? "Banner announcement updated and published successfully!"
                                        : "Banner announcement updated and scheduled for publication!";
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
            const oModel = this.getView().getModel("bannerModel");
            const bIsEditMode = oModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Reset to original values
                oModel.setProperty("/announcementContent", oModel.getProperty("/originalAnnouncementContent"));
                oModel.setProperty("/publishDate", oModel.getProperty("/originalPublishDate"));
                oModel.setProperty("/expiryDate", oModel.getProperty("/originalExpiryDate"));
                oModel.setProperty("/publishLater", oModel.getProperty("/originalPublishLater"));

                MessageToast.show("Reset to original values");
            } else {
                // Clear all fields
                this._initBannerModel();
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