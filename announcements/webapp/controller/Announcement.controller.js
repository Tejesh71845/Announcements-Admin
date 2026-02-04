sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/core/Fragment",
    "sap/ui/richtexteditor/RichTextEditor",
    "sap/ui/richtexteditor/library",
    "com/incture/announcements/utils/formatter"
], (Controller, JSONModel, MessageToast, MessageBox, Fragment, RTE, library, formatter) => {
    "use strict";

    return Controller.extend("com.incture.announcements.controller.Announcement", {

        /* ========================================
         * LIFECYCLE METHODS
         * ======================================== */

        onInit: function () {
            var oComponent = this.getOwnerComponent();
            this._router = oComponent.getRouter();
            this._router.getRoute("Announcement").attachPatternMatched(this._handleRouteMatched, this);
        },


        _handleRouteMatched: async function (oEvent) {
            //Initialise models and fields
            var oAnnouncementModel = this.getOwnerComponent().getModel("announcementModel");
            this._initWizardModel();
            this._initCategoryModel();
            this._initAnnouncementTypeModel();
            this._editContext = null;
            this._oWizardDialog = null;
            this._oRichTextEditor = null;

            var oAnnouncementsSmartTable = this.getView().byId("announcementsSmartTable");
            oAnnouncementsSmartTable.setModel(oAnnouncementModel);
        },

        onBeforeRebindTable: function (oEvent) {
            const oBindingParams = oEvent.getParameter("bindingParams");

            // Add the active filter
            const oActiveFilter = new sap.ui.model.Filter("isActive", sap.ui.model.FilterOperator.EQ, true);

            if (!oBindingParams.filters) {
                oBindingParams.filters = [];
            }
            oBindingParams.filters.push(oActiveFilter);

            // Ensure toTypes is expanded
            if (!oBindingParams.parameters) {
                oBindingParams.parameters = {};
            }
            oBindingParams.parameters.expand = "toTypes";
        },

        refreshSmartTable: function () {
            const oSmartTable = this.byId("announcementsSmartTable");
            const oModel = this.getOwnerComponent().getModel("announcementModel");

            console.log("Refreshing SmartTable...");

            // Step 1: Refresh the OData model
            if (oModel) {
                oModel.refresh(true); // Force refresh
                console.log("Model refreshed");
            }

            // Step 2: Refresh the SmartTable
            if (oSmartTable) {
                // Get the inner table
                const oTable = oSmartTable.getTable();

                if (oTable) {
                    // Refresh the table binding
                    const oBinding = oTable.getBinding("items");
                    if (oBinding) {
                        oBinding.refresh(true); // Force refresh
                        console.log("Table binding refreshed");
                    }
                }

                // Rebind the entire SmartTable
                oSmartTable.rebindTable();
                console.log("SmartTable rebound");
            }

            console.log("Complete refresh done");
        },

        onRefreshPress: function () {
            this.refreshSmartTable();
            MessageToast.show("Table refreshed");
        },

        /**
 * Check if an active Planned Scheduled announcement already exists
 * @returns {Promise} Promise that resolves with boolean (true if exists)
 */
        _checkActivePlannedScheduledExists: function () {
            return new Promise((resolve, reject) => {
                const oModel = this.getOwnerComponent().getModel("announcementModel");

                oModel.read("/Announcements", {
                    filters: [
                        new sap.ui.model.Filter("isActive", sap.ui.model.FilterOperator.EQ, true),
                        new sap.ui.model.Filter("announcementType", sap.ui.model.FilterOperator.Contains, "Planned Scheduled"),
                        new sap.ui.model.Filter({
                            filters: [
                                new sap.ui.model.Filter("announcementStatus", sap.ui.model.FilterOperator.EQ, "PUBLISHED"),
                                new sap.ui.model.Filter("announcementStatus", sap.ui.model.FilterOperator.EQ, "TO_BE_PUBLISHED")
                            ],
                            and: false
                        })
                    ],
                    success: (oData) => {
                        const bExists = oData.results && oData.results.length > 0;
                        resolve(bExists);
                    },
                    error: (oError) => {
                        console.error("Error checking Planned Scheduled announcements:", oError);
                        reject(oError);
                    }
                });
            });
        },

        /**
         * Check if the announcement being created/edited contains Planned Scheduled type
         * @param {Array} aAnnouncementTypeKeys - Array of selected announcement type keys
         * @param {string} sEditId - ID of announcement being edited (null for new)
         * @returns {boolean} True if Planned Scheduled is selected
         */
        _hasPlannedScheduledType: function (aAnnouncementTypeKeys, sEditId) {
            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];

            const aSelectedTypes = aAnnouncementTypeKeys
                .map(key => aTypeList.find(t => t.key === key)?.text)
                .filter(Boolean);

            return aSelectedTypes.some(type => type.toLowerCase().includes("planned scheduled"));
        },

        _initAnnouncementTypeModel: function () {
            const oAnnouncementTypeModel = new JSONModel();
            const sModelPath = sap.ui.require.toUrl("com/incture/announcements/model/AnnouncementTypes.json");

            oAnnouncementTypeModel.loadData(sModelPath);
            this.getView().setModel(oAnnouncementTypeModel, "announcementTypeModel");
        },

        // _initCategoryModel: function () {
        //     const oCategoryModel = new sap.ui.model.json.JSONModel({
        //         category: [],
        //         idToNameMap: {},
        //         nameToIdMap: {}
        //     });
        //     this.getView().setModel(oCategoryModel, "categoryModel");

        //     // Use destination with the correct path
        //     const sUrl = "/JnJ_Workzone_Portal_Destination_Node/odata/v2/type/Types";

        //     // Show busy indicator
        //     this.getView().setBusy(true);

        //     $.ajax({
        //         url: sUrl,
        //         method: "GET",
        //         dataType: "json",
        //         success: (oData) => {
        //             this.getView().setBusy(false);

        //             // OData V2 structure: data is in d.results
        //             const aEntries = oData.d?.results || [];
        //             const aDropdownData = [];
        //             const idToName = {};
        //             const nameToId = {};

        //             aEntries.forEach((entry) => {
        //                 // Use typeId and name from the response
        //                 const typeId = entry.typeId;
        //                 const typeName = entry.name;

        //                 // Validate data before processing
        //                 if (typeId && typeName) {
        //                     aDropdownData.push({
        //                         key: typeId,
        //                         text: typeName
        //                     });
        //                     idToName[typeId] = typeName;
        //                     nameToId[typeName] = typeId;
        //                 } else {
        //                     console.warn("Invalid type entry:", entry);
        //                 }
        //             });

        //             oCategoryModel.setProperty("/category", aDropdownData);
        //             oCategoryModel.setProperty("/idToNameMap", idToName);
        //             oCategoryModel.setProperty("/nameToIdMap", nameToId);

        //             console.log("Categories loaded:", aDropdownData.length);

        //             // Fetch announcements after categories are loaded
        //             this._fetchAnnouncements();
        //         },
        //         error: (jqXHR, textStatus, errorThrown) => {
        //             this.getView().setBusy(false);

        //             // Enhanced error logging
        //             console.error("Failed to fetch category data:", {
        //                 status: jqXHR.status,
        //                 statusText: textStatus,
        //                 error: errorThrown,
        //                 response: jqXHR.responseText
        //             });

        //             // Parse error response
        //             let errorMessage = "Failed to fetch category data";
        //             try {
        //                 if (jqXHR.responseText) {
        //                     const errorObj = JSON.parse(jqXHR.responseText);
        //                     errorMessage = errorObj.error?.message?.value || errorObj.error?.message || errorMessage;
        //                 }
        //             } catch (e) {
        //                 console.error("Error parsing error response:", e);
        //                 errorMessage = `${errorMessage} (Status: ${jqXHR.status})`;
        //             }

        //             // Show detailed error to user
        //             sap.m.MessageBox.error(errorMessage, {
        //                 title: "Error Loading Categories",
        //                 details: `${textStatus || "Internal Server Error"}\nStatus: ${jqXHR.status}`
        //             });
        //         }
        //     });
        // },


        _initCategoryModel: function () {
            const oCategoryModel = new sap.ui.model.json.JSONModel({
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

                    console.log("Categories loaded:", aDropdownData.length);

                    const oSmartTable = this.byId("announcementsSmartTable");
                    if (oSmartTable) {
                        oSmartTable.rebindTable();
                    }
                },
                error: (oError) => {
                    this.getView().setBusy(false);
                    console.error("Failed to fetch categories:", oError);
                    sap.m.MessageBox.error("Failed to load categories");
                }
            });
        },


        _initRichTextEditor: function (sContainerId) {
            if (this._oRichTextEditor) {
                this._oRichTextEditor.destroy();
            }

            const EditorType = library.EditorType;
            const oWizardModel = this.getView().getModel("wizardModel");
            const sDescription = oWizardModel.getProperty("/description") || "";

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

                    // Check character limit
                    const MAX_CHARS = 250;
                    if (sPlainText.length > MAX_CHARS) {
                        // Show warning message
                        if (!this._isCharLimitToastShown) {
                            MessageToast.show(`Description cannot exceed ${MAX_CHARS} characters. Current: ${sPlainText.length}`);
                            this._isCharLimitToastShown = true;

                            // Reset flag after 3 seconds to allow showing again
                            setTimeout(() => {
                                this._isCharLimitToastShown = false;
                            }, 3000);
                        }

                        // Revert to previous valid value
                        const sPreviousValue = oWizardModel.getProperty("/description") || "";
                        this._oRichTextEditor.setValue(sPreviousValue);
                        return;
                    }

                    // Reset toast flag on valid input
                    this._isCharLimitToastShown = false;

                    oWizardModel.setProperty("/description", sValue);
                    this._handleResetButtonVisibility(oWizardModel);
                    this._validateRichTextDescription(oWizardModel);
                }
            });

            const oContainer = this.byId(sContainerId);
            if (oContainer) {
                oContainer.removeAllItems();
                oContainer.addItem(this._oRichTextEditor);
            }
        },

        /* ========================================
         * DATA FETCHING
         * ======================================== */

        // _fetchAnnouncements: function () {
        //     const sUrl = "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/Announcements?$expand=toTypes($expand=type)";

        //     $.ajax({
        //         url: sUrl,
        //         method: "GET",
        //         dataType: "json",
        //         success: (oData) => {
        //             const oCategoryModel = this.getView().getModel("categoryModel");
        //             const idToNameMap = oCategoryModel.getProperty("/idToNameMap") || {};
        //             const aEntries = oData.value || [];

        //             // Filter out inactive announcements
        //             const aActiveEntries = aEntries.filter(entry => entry.isActive !== false);

        //             const aMapped = aActiveEntries.map((entry) => {
        //                 const aTypeIds = (entry.toTypes || [])
        //                     .map(item => item.type?.typeId)
        //                     .filter(Boolean);

        //                 const aCategoryNames = aTypeIds
        //                     .map(id => idToNameMap[id])
        //                     .filter(Boolean);

        //                 return {
        //                     id: entry.announcementId,
        //                     title: entry.title,
        //                     category: aCategoryNames.join(", ") || "N/A",
        //                     description: entry.description,
        //                     announcementType: entry.announcementType,
        //                     createdOn: entry.createdAt,
        //                     createdBy: entry.createdBy,
        //                     modifiedOn: entry.modifiedAt,
        //                     modifiedBy: entry.modifiedBy,
        //                     publishedAt: entry.publishedAt,           // NEW
        //                     publishedBy: entry.publishedBy,           // NEW
        //                     announcementStatus: entry.announcementStatus || "DRAFT",  // NEW
        //                     publishStatus: entry.publishStatus,       // NEW
        //                     startAnnouncement: entry.startAnnouncement, // NEW
        //                     endAnnouncement: entry.endAnnouncement,   // NEW
        //                     typeId: aTypeIds,
        //                     isActive: entry.isActive
        //                 };
        //             });

        //             // Sort announcements by createdOn date in descending order
        //             aMapped.sort((a, b) => {
        //                 const dateA = a.createdOn ? new Date(a.createdOn).getTime() : 0;
        //                 const dateB = b.createdOn ? new Date(b.createdOn).getTime() : 0;
        //                 return dateB - dateA;
        //             });

        //             this.getView().getModel().setProperty("/announcements", aMapped);
        //         },
        //         error: (xhr, status, err) => {
        //             console.error("Failed to fetch announcements:", status, err);
        //             sap.m.MessageBox.error("Unable to load announcements.");
        //         }
        //     });
        // },
        _initWizardModel: function () {
            const oWizardData = this._getDefaultWizardData();
            const oWizardModel = new JSONModel(oWizardData);
            this.getView().setModel(oWizardModel, "wizardModel");
        },

        _initWizardModelForEdit: function (oEditData) {
            // Convert announcementType string to array of keys
            let aAnnouncementTypeKeys = [];
            if (oEditData.announcementType) {
                const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
                const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];

                const aTypeTexts = oEditData.announcementType.split(",").map(t => t.trim()).filter(Boolean);
                aAnnouncementTypeKeys = aTypeTexts.map(text => {
                    const match = aTypeList.find(t => t.text.toLowerCase() === text.toLowerCase());
                    return match ? match.key : null;
                }).filter(Boolean);
            }

            // **FIX: Properly determine if publish data exists**
            const bHasPublishData = !!(oEditData.publishStatus && oEditData.endAnnouncement);

            const oWizardData = {
                ...this._getDefaultWizardData(),
                title: oEditData.title || "",
                announcementType: aAnnouncementTypeKeys,
                category: Array.isArray(oEditData.typeId) ? oEditData.typeId : [oEditData.typeId || ""],
                description: oEditData.description || "",
                originalTitle: oEditData.title || "",
                originalAnnouncementType: aAnnouncementTypeKeys,
                originalCategory: Array.isArray(oEditData.typeId) ? oEditData.typeId : [oEditData.typeId || ""],
                originalDescription: oEditData.description || "",
                currentFlow: "SINGLE",
                createdBy: oEditData.createdBy || this.getCurrentUserEmail(),
                createdOnFormatted: oEditData.createdOn
                    ? this.formatUSDateTime(oEditData.createdOn)
                    : this.formatUSDateTime(new Date()),

                // **FIX: Properly set publish option based on publishStatus**
                publishOption: oEditData.publishStatus === "PUBLISH_NOW" ? "PUBLISH" :
                    oEditData.publishStatus === "PUBLISH_LATER" ? "UNPUBLISH" : "",

                // **FIX: Format dates correctly for DatePicker**
                publishStartDate: oEditData.startAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.startAnnouncement)) : "",
                publishEndDate: oEditData.endAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.endAnnouncement)) : "",

                showDatePickers: bHasPublishData,
                startDateEnabled: oEditData.publishStatus === "PUBLISH_LATER",
                endDateEnabled: bHasPublishData,

                publishEnabled: bHasPublishData,
                cancelEnabled: true,

                // **FIX: Store original publish data for comparison**
                originalPublishOption: oEditData.publishStatus === "PUBLISH_NOW" ? "PUBLISH" :
                    oEditData.publishStatus === "PUBLISH_LATER" ? "UNPUBLISH" : "",
                originalPublishStartDate: oEditData.startAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.startAnnouncement)) : "",
                originalPublishEndDate: oEditData.endAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.endAnnouncement)) : "",

                isEditMode: true,
                editId: oEditData.id,
                showSingleWizard: true,
                showBulkWizard: false,
                singleCreateStepValidated: true,
                showUpdateButton: false,
                showReviewButton: true,
                showSubmitButton: false,
                showResetButton: false
            };

            const oWizardModel = new JSONModel(oWizardData);
            this.getView().setModel(oWizardModel, "wizardModel");
        },

        _initReviewRichTextEditor: function (sContainerId) {
            // Destroy existing review editor if any
            if (this._oReviewRichTextEditor) {
                this._oReviewRichTextEditor.destroy();
            }

            const EditorType = library.EditorType;
            const oWizardModel = this.getView().getModel("wizardModel");
            const sDescription = oWizardModel.getProperty("/description") || "";

            this._oReviewRichTextEditor = new RTE({
                editorType: EditorType.TinyMCE7,
                width: "100%",
                height: "300px",
                editable: false, // KEY: Make it read-only
                customToolbar: true,
                showGroupFont: true,
                showGroupLink: true,
                showGroupInsert: true,
                value: sDescription,
                ready: function () {
                    this.addButtonGroup("styles").addButtonGroup("table");
                },
            });

            const oContainer = this.byId(sContainerId);
            if (oContainer) {
                oContainer.removeAllItems();
                oContainer.addItem(this._oReviewRichTextEditor);
            }
        },

        _getDefaultWizardData: function () {
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            const oTomorrow = new Date(oToday);
            oTomorrow.setDate(oTomorrow.getDate() + 1);

            const oSevenDaysLater = new Date(oToday);
            oSevenDaysLater.setDate(oSevenDaysLater.getDate() + 30);

            return {
                // Field values
                title: "",
                announcementType: [],
                category: [],
                description: "",
                isDraftMode: false,

                // Publish Date fields
                publishOption: "",
                publishStartDate: "",
                publishEndDate: "",
                publishStartDateValueState: "None",
                publishStartDateValueStateText: "",
                publishEndDateValueState: "None",
                publishEndDateValueStateText: "",
                minPublishDate: oTomorrow,
                minEndDate: oToday,
                startDatePlaceholder: "Select start date",
                endDatePlaceholder: "Select end date",
                startDateEnabled: false,
                endDateEnabled: false,
                showDatePickers: false,

                // Status fields
                announcementStatus: "DRAFT",
                publishStatus: "",
                publishedBy: "",
                publishedAt: "",

                // Footer button enablement flags
                cancelEnabled: true,
                draftEnabled: true,
                submitEnabled: true,

                // Flow control
                currentFlow: "",
                showSingleWizard: true,
                showBulkWizard: false,

                // Step validation
                singleCreateStepValidated: false,
                bulkDownloadStepValidated: false,
                bulkUploadStepValidated: false,

                // Field validation states
                titleValueState: "None",
                titleValueStateText: "",
                announcementTypeValueState: "None",
                announcementTypeValueStateText: "",
                categoryValueState: "None",
                categoryValueStateText: "",
                descriptionValueState: "None",
                descriptionValueStateText: "",
                showValidationWarning: false,

                // Download/Upload status
                downloadStatus: "Pending",
                downloadMessage: "",
                downloadMessageType: "Information",
                showDownloadMessage: false,
                uploadStatus: "Pending",
                uploadMessage: "",
                uploadMessageType: "Information",
                showUploadMessage: false,
                templateDownloaded: "No",
                templateUploaded: "No",

                // FIXED: Button visibility - all hidden by default
                showResetButton: false,
                showSubmitButton: false,
                showDraftButton: false,
                showCancelButton: true, // Cancel always visible
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false,

                // Edit mode
                isEditMode: false,
                createdBy: this.getCurrentUserEmail(),
                createdOnFormatted: this.formatUSDateTime(new Date())
            };
        },

        _formatDateToDisplay: function (oDate) {
            return formatter.formatDateToDisplay(oDate);
        },


        formatStatusState: function (sStatus) {
            return formatter.formatStatusState(sStatus);
        },

        formatStatusText: function (sStatus) {
            return formatter.formatStatusText(sStatus);
        },


        /* ========================================
         * DIALOG MANAGEMENT
         * ======================================== */

        onCreatePress: function () {
            this._loadWizardDialog()
                .then(() => this._openWizard())
                .catch(error => {
                    MessageBox.error("Failed to load create dialog: " + error.message);
                });
        },

        _loadWizardDialog: function () {
            if (this._oWizardDialog) {
                return Promise.resolve();
            }

            return Fragment.load({
                id: this.getView().getId(),
                name: "com.incture.announcements.fragments.CreateAnnouncementWizard",
                type: "XML",
                controller: this
            }).then(oDialog => {
                this._oWizardDialog = oDialog;
                this.getView().addDependent(this._oWizardDialog);
            });
        },

        _openWizard: function () {
            if (!this._oWizardDialog || typeof this._oWizardDialog.open !== "function") {
                MessageBox.error("Dialog is not properly loaded. Please refresh and try again.");
                return;
            }

            if (!this._editContext) {
                this._initWizardModel();
                this._resetWizards();
                this._forceWizardRerender();
            }

            try {
                this._oWizardDialog.open();
            } catch (error) {
                MessageBox.error("Failed to open dialog: " + error.message);
            }
        },

        onCancelPress: function () {
            MessageBox.confirm(
                "Are you sure you want to cancel? All unsaved changes will be lost.",
                {
                    title: "Confirm Cancel",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.NO,
                    onClose: (oAction) => {
                        if (oAction === MessageBox.Action.YES) {
                            const oWizard = this.byId("singleWizard");
                            if (oWizard) {
                                oWizard.removeStyleClass("hideFirstWizardStep");
                            }

                            // CLEANUP BOTH RICHTEXTEDITORS
                            if (this._oRichTextEditor) {
                                this._oRichTextEditor.destroy();
                                this._oRichTextEditor = null;
                            }

                            if (this._oReviewRichTextEditor) {
                                this._oReviewRichTextEditor.destroy();
                                this._oReviewRichTextEditor = null;
                            }

                            this._editContext = null;
                            this._oWizardDialog.close();
                        }
                    }
                }
            );
        },

        /* ========================================
         * WIZARD MANAGEMENT
         * ======================================== */

        _resetWizards: function () {
            this._resetWizard("singleWizard", "singleSelectionStep");
            this._resetWizard("bulkWizard", "bulkSelectionStep");
        },

        _resetWizard: function (sWizardId, sStepId) {
            const oWizard = this.byId(sWizardId);
            const oStep = this.byId(sStepId);

            if (oWizard && oStep) {
                oWizard.discardProgress(oStep);
                oStep.setValidated(false);
            }
        },

        _forceWizardRerender: function () {
            ["singleWizard", "bulkWizard"].forEach(sWizardId => {
                const oWizard = this.byId(sWizardId);
                if (oWizard) oWizard.invalidate();
            });
        },

        _refreshWizardContent: function () {
            ["singleTitleInput", "singleAnnouncementTypeMultiComboBox", "singleCategoryMultiComboBox", "singleDescriptionTextArea"].forEach(sInputId => {
                const oInput = this.byId(sInputId);
                if (oInput) oInput.rerender();
            });
        },

        onWizardComplete: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/showSubmitButton", true);
        },

        /* ========================================
         * SINGLE WIZARD FLOW
         * ======================================== */

        onSingleFlowSelected: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/currentFlow", "SINGLE");

            // FIXED: Hide Draft and Submit buttons on selection step, keep Cancel visible
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });

            const oStep = this.byId("singleSelectionStep");
            if (oStep) oStep.setValidated(true);

            const oWizard = this.byId("singleWizard");
            if (oWizard) {
                oWizard.attachComplete(() => {
                    oWizardModel.setProperty("/showSubmitButton", true);
                });

                setTimeout(() => {
                    oWizard.nextStep();
                    this._refreshWizardContent();
                }, 200);
            }
        },
        onSwitchToBulkWizard: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/showSingleWizard", false);
            oWizardModel.setProperty("/showBulkWizard", true);
            oWizardModel.setProperty("/currentFlow", "BULK");

            // FIXED: Set button visibility for bulk flow
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });

            setTimeout(() => {
                this._forceWizardRerender();

                const oBulkWizard = this.byId("bulkWizard");
                const oBulkSelectionStep = this.byId("bulkSelectionStep");
                if (oBulkSelectionStep) oBulkSelectionStep.setValidated(true);
                if (oBulkWizard) {
                    setTimeout(() => {
                        oBulkWizard.nextStep();
                        this._refreshWizardContent();
                    }, 300);
                }
            }, 100);
        },

        onSingleCreateStepActivate: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");

            setTimeout(() => {
                this._initRichTextEditor("richTextContainer");
            }, bIsEditMode ? 200 : 100);

            // Show buttons for create/edit step - Reset button will be shown when user makes changes
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: true,
                showSubmitButton: true,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false,
                showResetButton: false // Will be shown when user makes changes via _handleResetButtonVisibility
            });
        },


        /**
 * Fetch CSRF Token
 */
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

        onSubmitPress: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");

            if (sCurrentFlow === "BULK") {
                const bulkData = oWizardModel.getProperty("/bulkData") || [];

                if (bulkData.length === 0) {
                    MessageBox.error("No data to submit. Please upload a valid Excel file.");
                    return;
                }

                const validationResult = this._validateBulkTableData(bulkData);
                if (!validationResult.isValid) {
                    this._showBulkValidationErrorDialog(validationResult.errors);
                    return;
                }

                // Check for Planned Scheduled in bulk data
                this._checkPlannedScheduledInBulk(bulkData);
            } else {
                // **FIX: Validate all fields and show errors**
                if (!this._validateAllFieldsWithErrors()) {
                    MessageToast.show("Please complete all required fields");
                    return;
                }

                if (!this._validatePublishOptions()) {
                    return;
                }

                // Check for Planned Scheduled announcement
                const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
                const sEditId = bIsEditMode ? oWizardModel.getProperty("/editId") : null;

                if (this._hasPlannedScheduledType(aAnnouncementTypeKeys, sEditId)) {
                    this._checkActivePlannedScheduledExists()
                        .then((bExists) => {
                            if (bExists && !bIsEditMode) {
                                MessageBox.error(
                                    "An active 'Planned Scheduled' announcement already exists in the system. " +
                                    "Only one active 'Planned Scheduled' announcement is allowed at a time. " +
                                    "Please delete or deactivate the existing one before creating a new one.",
                                    {
                                        title: "Duplicate Planned Scheduled Announcement"
                                    }
                                );
                                return;
                            }

                            // Proceed with submission
                            this._proceedWithSubmit(bIsEditMode);
                        })
                        .catch((oError) => {
                            MessageBox.error("Failed to validate announcement. Please try again.");
                        });
                } else {
                    // No Planned Scheduled type, proceed directly
                    this._proceedWithSubmit(bIsEditMode);
                }
            }
        },

        /**
         * Proceed with announcement submission after validation
         */
        _proceedWithSubmit: function (bIsEditMode) {
            const sConfirmMessage = bIsEditMode
                ? "Are you sure you want to submit these changes?"
                : "Are you sure you want to submit this announcement?";

            MessageBox.confirm(sConfirmMessage, {
                title: "Confirm Submit",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.YES,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        if (bIsEditMode) {
                            this._handleEditSubmit();
                        } else {
                            this._handleSubmitAnnouncement();
                        }
                    }
                }
            });
        },

        _validateAllFieldsWithErrors: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            // const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();

            let bValid = true;

            // Validate Title
            if (!sTitle) {
                this._setValidationState(oWizardModel, "title", false, "Title is required");
                bValid = false;
            } else {
                this._setValidationState(oWizardModel, "title", true, "");
            }

            // Validate Announcement Type
            if (aAnnouncementType.length === 0) {
                this._setValidationState(oWizardModel, "announcementType", false, "At least one announcement type is required");
                bValid = false;
            } else {
                this._setValidationState(oWizardModel, "announcementType", true, "");
            }

            // Validate Category
            // if (aCategory.length === 0) {
            //     this._setValidationState(oWizardModel, "category", false, "At least one category is required");
            //     bValid = false;
            // } else {
            //     this._setValidationState(oWizardModel, "category", true, "");
            // }

            this._setValidationState(oWizardModel, "category", true, "");

            // Validate Description (RichTextEditor)
            if (!sPlainText) {
                oWizardModel.setProperty("/descriptionValueState", "Error");
                oWizardModel.setProperty("/descriptionValueStateText", "Description is required");

                // Add red border to RichTextEditor container
                const oContainer = this.byId("richTextContainer");
                if (oContainer) {
                    oContainer.addStyleClass("richTextError");
                }
                bValid = false;
            } else {
                oWizardModel.setProperty("/descriptionValueState", "None");
                oWizardModel.setProperty("/descriptionValueStateText", "");

                // Remove red border
                const oContainer = this.byId("richTextContainer");
                if (oContainer) {
                    oContainer.removeStyleClass("richTextError");
                }
            }

            return bValid;
        },

        _validateMandatoryFields: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            // const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();

            return sTitle && aAnnouncementType.length > 0 && sDescription;
        },

        _handleSubmitAnnouncement: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aTypeIds = oWizardModel.getProperty("/category") || [];
            const sPublishOption = oWizardModel.getProperty("/publishOption");
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            const sEndDate = oWizardModel.getProperty("/publishEndDate");

            const sDescription = this._stripHtmlTags(sDescriptionHTML);

            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];
            const aSelectedTypes = aAnnouncementTypeKeys
                .map(key => aTypeList.find(t => t.key === key)?.text)
                .filter(Boolean);
            const sAnnouncementType = aSelectedTypes.join(",");

            const oBusy = new sap.m.BusyDialog({ text: "Submitting announcement..." });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement;
                    const currentDateTime = new Date().toISOString();

                    if (sPublishOption === "PUBLISH") {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sStartDate).toISOString();
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    }

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
                            toTypes: aTypeIds.map(typeId => ({
                                type: { typeId: typeId }
                            }))
                        }]
                    };

                    // Fetch CSRF token first
                    this._getCSRFToken()
                        .then((csrfToken) => {
                            $.ajax({
                                url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/bulkCreateAnnouncements",
                                method: "POST",
                                contentType: "application/json",
                                dataType: "json",
                                headers: {
                                    "X-CSRF-Token": csrfToken  // Include CSRF token
                                },
                                data: JSON.stringify(oPayload),
                                success: (oResponse) => {
                                    oBusy.close();
                                    this._oWizardDialog.close();
                                    const sMessage = announcementStatus === "PUBLISHED"
                                        ? `Announcement '${sTitle}' published successfully!`
                                        : `Announcement '${sTitle}' scheduled for publication!`;
                                    sap.m.MessageToast.show(sMessage);

                                    // Add delay before refresh
                                    setTimeout(() => {
                                        this.refreshSmartTable();
                                    }, 500);
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Create announcement failed:", status, err);
                                    console.error("Response:", xhr.responseText);
                                    let sErrorMessage = "Failed to create announcement. Please try again.";
                                    if (xhr.responseJSON?.error?.message) {
                                        sErrorMessage = xhr.responseJSON.error.message;
                                    }
                                    sap.m.MessageBox.error(sErrorMessage);
                                }
                            });
                        })
                        .catch((err) => {
                            oBusy.close();
                            console.error("CSRF token fetch failed:", err);
                            sap.m.MessageBox.error("Failed to initialize request. Please try again.");
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    sap.m.MessageBox.error("Failed to get current user: " + error.message);
                });
        },
        /* ========================================
         * INPUT VALIDATION & CHANGE HANDLERS
         * ======================================== */

        _validateRichTextDescription: function (oWizardModel) {
            const sDescriptionHTML = oWizardModel.getProperty("/description") || "";
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();
            const bValid = sPlainText.length > 0;

            // Store validation state in wizard model (not in RTE)
            oWizardModel.setProperty("/descriptionValueState", bValid ? "None" : "Error");
            oWizardModel.setProperty("/descriptionValueStateText", bValid ? "" : "Description is required");

            // Visual feedback using CSS class instead of valueState
            const oContainer = this.byId("richTextContainer");
            if (oContainer) {
                if (!bValid) {
                    oContainer.addStyleClass("richTextError");
                } else {
                    oContainer.removeStyleClass("richTextError");
                }
            }

            // Check overall step validity
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            // const aCategory = oWizardModel.getProperty("/category") || [];

            const bOverallValid = sTitle.length > 0 &&
                aAnnouncementType.length > 0 &&
                bValid;

            oWizardModel.setProperty("/singleCreateStepValidated", bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) {
                oStep.setValidated(bOverallValid);
            }
        },

        onSingleInputChange: function (oEvent) {
            const oSource = oEvent.getSource();
            const sControlType = oSource.getMetadata().getName();
            const oWizardModel = this.getView().getModel("wizardModel");
            let sValue = oEvent.getParameter("value") || "";

            if (sControlType === "sap.m.Input" || sControlType === "sap.m.TextArea") {
                // Remove special characters (keep your existing logic)
                sValue = sValue.replace(/[^a-zA-Z0-9 ]/g, "");

                // Get max length from control or use default
                const iMaxLength = oSource.getMaxLength && oSource.getMaxLength() > 0
                    ? oSource.getMaxLength()
                    : 100;

                // **ENFORCE HARD LIMIT**
                if (sValue.length > iMaxLength) {
                    sValue = sValue.substring(0, iMaxLength);
                    oSource.setValue(sValue); // Update the input field immediately
                }

                // Update value if it was modified
                if (sValue !== oEvent.getParameter("value")) {
                    oSource.setValue(sValue);
                }
            }

            this._updateInputValue(oSource, sValue, oWizardModel);
            this._handleResetButtonVisibility(oWizardModel);
            this._validateField(oSource, oWizardModel);
        },

        onSingleAnnouncementTypeSelectionChange: function (oEvent) {
            const aSelectedKeys = oEvent.getSource().getSelectedKeys();
            const aSelectedItems = oEvent.getSource().getSelectedItems();
            const aTypeTexts = aSelectedItems.map((item) => item.getText());

            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/announcementType", aSelectedKeys);
            oWizardModel.setProperty("/announcementTypeDisplayName", aTypeTexts.join(", "));

            this._validateMultiAnnouncementType(oWizardModel);
            this._handleResetButtonVisibility(oWizardModel);
        },

        onSingleCategorySelectionChange: function (oEvent) {
            const aSelectedKeys = oEvent.getSource().getSelectedKeys();
            const aSelectedItems = oEvent.getSource().getSelectedItems();
            const aCategoryTexts = aSelectedItems.map((item) => item.getText());

            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/category", aSelectedKeys);
            oWizardModel.setProperty("/categoryDisplayName", aCategoryTexts.join(", "));

            this._validateMultiCategory(oWizardModel);
            this._handleResetButtonVisibility(oWizardModel);
        },

        _updateInputValue: function (oSource, sValue, oWizardModel) {
            const sId = oSource.getId();

            if (sId.indexOf("singleTitleInput") > -1) {
                oWizardModel.setProperty("/title", sValue);
            } else if (sId.indexOf("singleDescriptionTextArea") > -1) {
                oWizardModel.setProperty("/description", sValue);
            }
        },

        _validateField: function (oSource, oWizardModel) {
            const sId = oSource.getId();
            let sFieldKey = "";
            let sErrorMessage = "";
            let sValue = "";

            switch (sId) {
                case this.createId("singleTitleInput"):
                    sFieldKey = "title";
                    sErrorMessage = "Title is required";
                    sValue = (oWizardModel.getProperty("/title") || "").trim();
                    break;

                case this.createId("singleDescriptionTextArea"):
                    sFieldKey = "description";
                    sErrorMessage = "Description is required";
                    sValue = (oWizardModel.getProperty("/description") || "").trim();
                    break;

                default:
                    return;
            }

            const bValid = sValue.length > 0;
            this._setValidationState(oWizardModel, sFieldKey, bValid, sErrorMessage);

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();

            const bOverallValid = sTitle && aAnnouncementType.length > 0 && sDescription;

            oWizardModel.setProperty("/singleCreateStepValidated", !!bOverallValid);
            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(!!bOverallValid);
        },

        _setValidationState: function (oWizardModel, sField, bValid, sErrorMessage) {
            const sState = bValid ? "None" : "Error";
            const sMessage = bValid ? "" : sErrorMessage;

            oWizardModel.setProperty(`/${sField}ValueState`, sState);
            oWizardModel.setProperty(`/${sField}ValueStateText`, sMessage);
        },

        _validateMultiAnnouncementType: function (oWizardModel) {
            const aTypes = oWizardModel.getProperty("/announcementType") || [];
            const bValid = Array.isArray(aTypes) && aTypes.length > 0;
            this._setValidationState(oWizardModel, "announcementType", bValid, "At least one announcement type is required");

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();
            const bOverallValid = sTitle && bValid && sDescription;
            oWizardModel.setProperty("/singleCreateStepValidated", !!bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(!!bOverallValid);
        },

        _validateMultiCategory: function (oWizardModel) {
            const aCategories = oWizardModel.getProperty("/category") || [];
            // const bValid = Array.isArray(aCategories) && aCategories.length > 0;
            // this._setValidationState(oWizardModel, "category", bValid, "At least one category is required");

            this._setValidationState(oWizardModel, "category", true, "");

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();
            const bOverallValid = sTitle && aAnnouncementType.length > 0 && sDescription;
            oWizardModel.setProperty("/singleCreateStepValidated", !!bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(!!bOverallValid);
        },

        _handleResetButtonVisibility: function (oWizardModel) {
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");

            // Only show Reset button for SINGLE flow
            if (sCurrentFlow !== "SINGLE") {
                oWizardModel.setProperty("/showResetButton", false);
                return;
            }

            const sTitle = oWizardModel.getProperty("/title") || "";
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = oWizardModel.getProperty("/description") || "";

            // ✅ FIX: Also check publish-related fields
            const sPublishOption = oWizardModel.getProperty("/publishOption") || "";
            const sPublishStartDate = oWizardModel.getProperty("/publishStartDate") || "";
            const sPublishEndDate = oWizardModel.getProperty("/publishEndDate") || "";

            if (bIsEditMode) {
                // Edit Mode: Show Reset if any field differs from original
                const sOriginalTitle = oWizardModel.getProperty("/originalTitle") || "";
                const aOriginalAnnouncementType = oWizardModel.getProperty("/originalAnnouncementType") || [];
                const aOriginalCategory = oWizardModel.getProperty("/originalCategory") || [];
                const sOriginalDescription = oWizardModel.getProperty("/originalDescription") || "";

                // ✅ FIX: Also compare publish-related original fields
                const sOriginalPublishOption = oWizardModel.getProperty("/originalPublishOption") || "";
                const sOriginalPublishStartDate = oWizardModel.getProperty("/originalPublishStartDate") || "";
                const sOriginalPublishEndDate = oWizardModel.getProperty("/originalPublishEndDate") || "";

                const bHasChanged = (sTitle !== sOriginalTitle) ||
                    (JSON.stringify([...aAnnouncementType].sort()) !== JSON.stringify([...aOriginalAnnouncementType].sort())) ||
                    (JSON.stringify([...aCategory].sort()) !== JSON.stringify([...aOriginalCategory].sort())) ||
                    (sDescription !== sOriginalDescription) ||
                    (sPublishOption !== sOriginalPublishOption) ||
                    (sPublishStartDate !== sOriginalPublishStartDate) ||
                    (sPublishEndDate !== sOriginalPublishEndDate);

                oWizardModel.setProperty("/showResetButton", bHasChanged);
            } else {
                // Create Mode: Show Reset if user has entered any data
                const bUserHasStartedTyping = sTitle.length > 0 ||
                    aAnnouncementType.length > 0 ||
                    aCategory.length > 0 ||
                    sDescription.length > 0 ||
                    sPublishOption.length > 0 ||
                    sPublishStartDate.length > 0 ||
                    sPublishEndDate.length > 0;

                oWizardModel.setProperty("/showResetButton", bUserHasStartedTyping);
            }
        },

        _clearValidationErrors: function (oWizardModel) {
            oWizardModel.setProperty("/titleValueState", "None");
            oWizardModel.setProperty("/titleValueStateText", "");
            oWizardModel.setProperty("/announcementTypeValueState", "None");
            oWizardModel.setProperty("/announcementTypeValueStateText", "");
            oWizardModel.setProperty("/categoryValueState", "None");
            oWizardModel.setProperty("/categoryValueStateText", "");
            oWizardModel.setProperty("/descriptionValueState", "None");
            oWizardModel.setProperty("/descriptionValueStateText", "");
            oWizardModel.setProperty("/showValidationWarning", false);
        },

        /* ========================================
         * REVIEW & UPDATE HANDLERS
         * ======================================== */


        onReviewPress: function () {
            if (!this._validateSingleCreateStep()) {
                MessageToast.show("Please complete all required fields with valid information");
                return;
            }

            const oWizardModel = this.getView().getModel("wizardModel");
            const oCategoryModel = this.getView().getModel("categoryModel");
            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");

            // Set announcement type display name for review
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aAnnouncementTypes = oAnnouncementTypeModel.getProperty("/types") || [];
            const aAnnouncementTypeNames = aAnnouncementTypeKeys
                .map(key => aAnnouncementTypes.find(t => t.key === key)?.text)
                .filter(Boolean);
            oWizardModel.setProperty("/announcementTypeDisplayName", aAnnouncementTypeNames.join(", "));

            // Set category display name for review
            const aTypeIds = oWizardModel.getProperty("/category") || [];
            const idToNameMap = oCategoryModel.getProperty("/idToNameMap");
            const aCategoryNames = aTypeIds.map((id) => idToNameMap[id] || id);
            oWizardModel.setProperty("/categoryDisplayName", aCategoryNames.join(", "));

            // Reset publish date fields
            // oWizardModel.setProperty("/publishOption", "");
            // oWizardModel.setProperty("/publishStartDate", "");
            // oWizardModel.setProperty("/publishEndDate", "");
            // oWizardModel.setProperty("/showDatePickers", false);
            // oWizardModel.setProperty("/startDateEnabled", false);
            // oWizardModel.setProperty("/endDateEnabled", false);

            const bIsEditMode = oWizardModel.getProperty("/isEditMode");

            if (!bIsEditMode) {
                oWizardModel.setProperty("/publishOption", "");
                oWizardModel.setProperty("/publishStartDate", "");
                oWizardModel.setProperty("/publishEndDate", "");
                oWizardModel.setProperty("/showDatePickers", false);
                oWizardModel.setProperty("/startDateEnabled", false);
                oWizardModel.setProperty("/endDateEnabled", false);
            } else {
                // In edit mode, ensure date pickers are shown if publish option exists
                const sPublishOption = oWizardModel.getProperty("/publishOption");
                if (sPublishOption) {
                    oWizardModel.setProperty("/showDatePickers", true);
                    oWizardModel.setProperty("/startDateEnabled", sPublishOption === "UNPUBLISH");
                    oWizardModel.setProperty("/endDateEnabled", true);

                    const sEndDate = oWizardModel.getProperty("/publishEndDate");
                    if (sEndDate) {
                        oWizardModel.setProperty("/publishEnabled", true);
                    }
                }
            }

            // Update button visibility based on mode
            if (bIsEditMode) {
                this._setButtonVisibility({
                    showReviewButton: false,
                    showPublishButton: true,
                    showDraftButton: true,
                    showUpdateButton: false,
                    showSubmitButton: false,
                    showResetButton: false,
                    showCancelButton: true
                });
            } else {
                this._setButtonVisibility({
                    showReviewButton: false,
                    showPublishButton: true,
                    showDraftButton: true,
                    showSubmitButton: false,
                    showUpdateButton: false,
                    showResetButton: false,
                    showCancelButton: true
                });
            }

            const oWizard = this.byId("singleWizard");
            const oReviewStep = this.byId("singleReviewStep");

            if (oWizard && oReviewStep) {
                // Validate the create step before proceeding
                const oCreateStep = this.byId("singleCreateStep");
                if (oCreateStep) {
                    oCreateStep.setValidated(true);
                }

                // Navigate to review step
                oWizard.nextStep();

                // Initialize review RichTextEditor after navigation
                setTimeout(() => {
                    this._initReviewRichTextEditor("reviewRichTextContainer");
                }, 200);
            }
        },

        /**
         * RadioButton select handler for Publish / Unpublish
         */
        onPublishOptionRadioSelect: function (oEvent) {
            const oSource = oEvent.getSource();
            const sId = oSource.getId();
            const oWizardModel = this.getView().getModel("wizardModel");

            let sOption = "";

            if (sId.indexOf("publishPublishRadio") > -1) {
                sOption = "PUBLISH";
            } else if (sId.indexOf("publishUnpublishRadio") > -1) {
                sOption = "UNPUBLISH";
            }

            if (!sOption) {
                return;
            }

            oWizardModel.setProperty("/publishOption", sOption);
            oWizardModel.setProperty("/showDatePickers", true);

            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (sOption === "PUBLISH") {
                const oEndDate = new Date(oToday);
                oEndDate.setDate(oEndDate.getDate() + 30);

                const sTodayValue = formatter.formatDateToValue(oToday);
                const sEndDateValue = formatter.formatDateToValue(oEndDate);

                oWizardModel.setProperty("/startDateEnabled", false);
                oWizardModel.setProperty("/endDateEnabled", true);
                oWizardModel.setProperty("/publishStartDate", sTodayValue);
                oWizardModel.setProperty("/publishEndDate", sEndDateValue);
                oWizardModel.setProperty("/publishStartDateValueState", "None");
                oWizardModel.setProperty("/publishStartDateValueStateText", "");
                oWizardModel.setProperty("/publishEndDateValueState", "None");
                oWizardModel.setProperty("/publishEndDateValueStateText", "");

                // **CHANGED: Set minEndDate to day AFTER today for Publish Now**
                const oMinEndDate = new Date(oToday);
                oMinEndDate.setDate(oMinEndDate.getDate() + 1);
                oWizardModel.setProperty("/minEndDate", oMinEndDate);

                oWizardModel.setProperty("/draftEnabled", true);
            } else {
                const oTomorrow = new Date(oToday);
                oTomorrow.setDate(oTomorrow.getDate() + 1);
                const oEndDate = new Date(oTomorrow);
                oEndDate.setDate(oEndDate.getDate() + 30);

                const sTomorrowValue = formatter.formatDateToValue(oTomorrow);
                const sEndDateValue = formatter.formatDateToValue(oEndDate);

                oWizardModel.setProperty("/startDateEnabled", true);
                oWizardModel.setProperty("/endDateEnabled", true);
                oWizardModel.setProperty("/publishStartDate", sTomorrowValue);
                oWizardModel.setProperty("/publishEndDate", sEndDateValue);
                oWizardModel.setProperty("/minPublishDate", oTomorrow);

                // **CHANGED: Set minEndDate to day AFTER tomorrow for Publish Later**
                const oMinEndDate = new Date(oTomorrow);
                oMinEndDate.setDate(oMinEndDate.getDate() + 1);
                oWizardModel.setProperty("/minEndDate", oMinEndDate);

                oWizardModel.setProperty("/draftEnabled", false);
            }

            oWizardModel.setProperty("/cancelEnabled", true);
            oWizardModel.setProperty("/publishEnabled", true);

            // Check if Reset button should be visible after publish option change
            this._handleResetButtonVisibility(oWizardModel);
        },
        _formatDateToValue: function (oDate) {
            return formatter.formatDateToValue(oDate);
        },

        /**
         * DatePicker change handler for Start Date (Unpublish)
         */
        onPublishStartDateChange: function (oEvent) {
            const sValue = oEvent.getParameter("value");
            const oWizardModel = this.getView().getModel("wizardModel");

            oWizardModel.setProperty("/publishStartDate", sValue);

            if (!sValue) {
                oWizardModel.setProperty("/publishStartDateValueState", "Error");
                oWizardModel.setProperty("/publishStartDateValueStateText", "Start Date is required");
                oWizardModel.setProperty("/publishEndDate", "");
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);
            const oTomorrow = new Date(oToday);
            oTomorrow.setDate(oTomorrow.getDate() + 1);

            // For Publish Later, start date must be tomorrow or later
            if (oSelected < oTomorrow) {
                oWizardModel.setProperty("/publishStartDateValueState", "Error");
                oWizardModel.setProperty("/publishStartDateValueStateText", "Start Date must be tomorrow or later for Publish Later");
                oWizardModel.setProperty("/publishEndDate", "");
                return;
            }

            oWizardModel.setProperty("/publishStartDateValueState", "None");
            oWizardModel.setProperty("/publishStartDateValueStateText", "");

            // Auto-populate End Date (Start Date + 30 days)
            const oEndDate = new Date(oSelected);
            oEndDate.setDate(oEndDate.getDate() + 30);
            const sEndDateValue = this._formatDateToValue(oEndDate);

            oWizardModel.setProperty("/publishEndDate", sEndDateValue);
            oWizardModel.setProperty("/publishEndDateValueState", "None");
            oWizardModel.setProperty("/publishEndDateValueStateText", "");

            // **CHANGED: Update minimum date for End Date picker to day AFTER selected start date**
            const oMinEndDate = new Date(oSelected);
            oMinEndDate.setDate(oMinEndDate.getDate() + 1);
            oWizardModel.setProperty("/minEndDate", oMinEndDate);

            this._handleResetButtonVisibility(oWizardModel);
        },
        onPublishEndDateChange: function (oEvent) {
            const sValue = oEvent.getParameter("value");
            const oWizardModel = this.getView().getModel("wizardModel");

            oWizardModel.setProperty("/publishEndDate", sValue);

            if (!sValue) {
                oWizardModel.setProperty("/publishEndDateValueState", "Error");
                oWizardModel.setProperty("/publishEndDateValueStateText", "End Date is required");
                return;
            }

            const oSelected = new Date(sValue);
            oSelected.setHours(0, 0, 0, 0);
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            if (oSelected < oToday) {
                oWizardModel.setProperty("/publishEndDateValueState", "Error");
                oWizardModel.setProperty("/publishEndDateValueStateText", "End Date cannot be in the past");
                return;
            }

            // **CHANGED: Validate End Date is AFTER Start Date (not equal)**
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            if (sStartDate) {
                const oStartDate = new Date(sStartDate);
                oStartDate.setHours(0, 0, 0, 0);
                if (oSelected <= oStartDate) { // **CHANGED: from < to <=**
                    oWizardModel.setProperty("/publishEndDateValueState", "Error");
                    oWizardModel.setProperty("/publishEndDateValueStateText", "End Date must be after Start Date");
                    return;
                }
            }

            oWizardModel.setProperty("/publishEndDateValueState", "None");
            oWizardModel.setProperty("/publishEndDateValueStateText", "");

            this._handleResetButtonVisibility(oWizardModel);
        },

        _validateSingleCreateStep: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            // const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();

            const bTitleValid = sTitle.length > 0;
            const bAnnouncementTypeValid = aAnnouncementType.length > 0;
            // const bCategoryValid = aCategory.length > 0;
            const bDescriptionValid = sDescription.length > 0;

            const bOverallValid = bTitleValid && bAnnouncementTypeValid && bDescriptionValid;

            this._setValidationState(oWizardModel, "title", bTitleValid, "Title is required");
            this._setValidationState(oWizardModel, "announcementType", bAnnouncementTypeValid, "Announcement Type is required");
            // this._setValidationState(oWizardModel, "category", bCategoryValid, "Category is required");
            this._setValidationState(oWizardModel, "description", bDescriptionValid, "Description is required");

            oWizardModel.setProperty("/singleCreateStepValidated", bOverallValid);
            oWizardModel.setProperty("/showValidationWarning", !bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(bOverallValid);

            return bOverallValid;
        },

        /* ========================================
         * BULK WIZARD FLOW
         * ======================================== */

        onBulkFlowSelected: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/currentFlow", "BULK");

            // FIXED: Hide all buttons except Cancel on selection step
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });

            const oStep = this.byId("bulkSelectionStep");
            if (oStep) oStep.setValidated(true);

            const oWizard = this.byId("bulkWizard");
            if (oWizard) {
                oWizard.attachComplete(() => {
                    // Submit button will be shown in Review step
                });

                setTimeout(() => {
                    oWizard.nextStep();
                    this._refreshWizardContent();
                }, 200);
            }
        },

        onBulkDownloadStepActivate: function () {
            const oWizardModel = this.getView().getModel("wizardModel");

            // Only Cancel button visible
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });
        },

        onBulkUploadStepActivate: function () {
            const oWizardModel = this.getView().getModel("wizardModel");

            // Only Cancel button visible
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });
        },

        onBulkReviewStepActivate: function () {
            const oWizardModel = this.getView().getModel("wizardModel");

            // FIXED: Show Submit button only on Review step
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: true,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });

            oWizardModel.setProperty("/submitEnabled", true);
        },
        onSwitchToSingleWizard: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/showSingleWizard", true);
            oWizardModel.setProperty("/showBulkWizard", false);
            oWizardModel.setProperty("/currentFlow", "SINGLE");

            // FIXED: Set button visibility for single flow
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: false,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false
            });

            setTimeout(() => {
                this._forceWizardRerender();

                const oSingleWizard = this.byId("singleWizard");
                const oSingleSelectionStep = this.byId("singleSelectionStep");
                if (oSingleSelectionStep) oSingleSelectionStep.setValidated(true);
                if (oSingleWizard) {
                    setTimeout(() => {
                        oSingleWizard.nextStep();
                        this._refreshWizardContent();
                    }, 300);
                }
            }, 100);
        },

        /* ========================================
         * BULK TEMPLATE DOWNLOAD
         * ======================================== */

        onBulkTableInputChange: function (oEvent) {
            // Mark that data has been modified
            const oWizardModel = this.getView().getModel("wizardModel");
            oWizardModel.setProperty("/bulkDataModified", true);
        },

        onDownloadTemplatePress: async function () {
            const oModel = this.getView().getModel("wizardModel");
            const componentId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
            const componentPath = componentId.replace(/\./g, "/");
            const oPath = sap.ui.require.toUrl(componentPath + "/artifacts") + "/Announcements_Template.xlsx";

            try {
                const response = await fetch(oPath, { cache: "no-store" });
                if (!response.ok) throw new Error("HTTP error " + response.status);

                const blob = await response.blob();
                this._downloadBlob(blob, "Announcements_Template.xlsx");

                this._setDownloadStatus(oModel, "Downloaded", "Template downloaded successfully.", "Success");
                MessageToast.show("Template downloaded successfully");

                const oBulkWizard = this.byId("bulkWizard");
                if (oBulkWizard) {
                    setTimeout(() => oBulkWizard.nextStep(), 1000);
                }
            } catch (err) {
                console.error("Download failed:", err);
                this._setDownloadStatus(oModel, "Failed", "Error occurred while downloading template.", "Error");
                MessageToast.show("Error downloading template: " + err.message);
            }
        },

        _downloadBlob: function (blob, fileName) {
            const link = document.createElement("a");
            const url = window.URL.createObjectURL(blob);
            link.href = url;
            link.download = fileName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            setTimeout(() => window.URL.revokeObjectURL(url), 10);
        },

        _setDownloadStatus: function (oModel, sStatus, sMessage, sType) {
            oModel.setProperty("/downloadStatus", sStatus);
            oModel.setProperty("/downloadMessage", sMessage);
            oModel.setProperty("/downloadMessageType", sType);
            oModel.setProperty("/showDownloadMessage", true);
        },

        /* ========================================
         * BULK TEMPLATE UPLOAD
         * ======================================== */



        onUploadTemplatePress: function () {
            const fileInput = document.createElement("input");
            fileInput.type = "file";
            fileInput.accept = ".xlsx, .xls";

            fileInput.onchange = (event) => {
                const file = event.target.files[0];
                if (!file) return;

                const fileExtension = file.name.split(".").pop().toLowerCase();
                if (fileExtension === "xlsx" || fileExtension === "xls") {
                    this._handleExcelUpload(file);
                } else {
                    MessageBox.error("Please upload a valid Excel file (.xlsx or .xls)");
                }
            };

            fileInput.click();
        },

        _handleExcelUpload: function (file) {
            const oWizardModel = this.getView().getModel("wizardModel");
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const structureValidation = this._validateExcelStructure(e.target.result);

                    if (!structureValidation.isValid) {
                        MessageBox.error(structureValidation.errorMessage, {
                            title: "Invalid Template Structure",
                            contentWidth: "500px"
                        });
                        this._setUploadStatus(oWizardModel, "Failed", "Invalid template structure.", "Error");
                        return;
                    }

                    const result = this._parseExcelDataWithoutValidation(e.target.result);

                    if (result.totalRows === 0) {
                        MessageBox.warning("The uploaded Excel file contains no data. Please add at least one row.");
                        this._setUploadStatus(oWizardModel, "Failed", "No data found in Excel file.", "Warning");
                        return;
                    }

                    // CHANGED: Don't validate during upload, just parse the data
                    this._handleSuccessfulUpload(oWizardModel, result.data);
                } catch (error) {
                    this._handleUploadError(oWizardModel, error);
                }
            };

            reader.readAsArrayBuffer(file);
        },

        _parseExcelDataWithoutValidation: function (arrayBuffer) {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: ["title", "announcementType", "category", "description", "startDate", "endDate"],
                range: 1,
                raw: false,
                dateNF: "dd/mm/yyyy"
            });

            const parsedData = [];

            jsonData.forEach((row, index) => {
                const title = row.title ? row.title.toString().trim() : "";
                const announcementType = row.announcementType ? row.announcementType.toString().trim() : "";
                const category = row.category ? row.category.toString().trim() : "";
                const description = row.description ? row.description.toString().trim() : "";
                const startDate = row.startDate ? row.startDate.toString().trim() : "";
                const endDate = row.endDate ? row.endDate.toString().trim() : "";

                // Parse dates if available
                let formattedStartDate = "";
                let formattedEndDate = "";

                if (startDate) {
                    const parsedStart = this._parseExcelDate(startDate);
                    if (parsedStart) {
                        formattedStartDate = this._formatDateToValue(parsedStart);
                    }
                }

                if (endDate) {
                    const parsedEnd = this._parseExcelDate(endDate);
                    if (parsedEnd) {
                        formattedEndDate = this._formatDateToValue(parsedEnd);
                    }
                }

                parsedData.push({
                    title: title,
                    announcementType: announcementType,
                    category: category,
                    description: description,
                    startDate: formattedStartDate,
                    endDate: formattedEndDate
                });
            });

            return {
                data: parsedData,
                totalRows: jsonData.length
            };
        },

        _validateExcelStructure: function (arrayBuffer) {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const expectedHeaders = ["Title", "Announcement Type", "Category", "Description", "Start Date", "End Date"];

            const range = XLSX.utils.decode_range(worksheet["!ref"]);
            const actualHeaders = [];

            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                const cell = worksheet[cellAddress];
                if (cell && cell.v) {
                    actualHeaders.push(cell.v.toString().trim());
                }
            }

            if (actualHeaders.length !== expectedHeaders.length) {
                return {
                    isValid: false,
                    errorMessage: `Invalid template structure!\n\nExpected ${expectedHeaders.length} columns but found ${actualHeaders.length} columns.\n\nExpected columns:\n${expectedHeaders.join(", ")}\n\nFound columns:\n${actualHeaders.join(", ")}\n\nPlease download the template again and do not modify the column structure.`
                };
            }

            for (let i = 0; i < expectedHeaders.length; i++) {
                if (actualHeaders[i] !== expectedHeaders[i]) {
                    return {
                        isValid: false,
                        errorMessage: `Invalid template structure!\n\nColumn mismatch at position ${i + 1}:\nExpected: "${expectedHeaders[i]}"\nFound: "${actualHeaders[i]}"\n\nExpected columns:\n${expectedHeaders.join(", ")}\n\nPlease download the template again and do not modify the column headers.`
                    };
                }
            }

            return { isValid: true };
        },

        _parseExcelData: function (arrayBuffer) {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array", cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: ["title", "announcementType", "category", "description", "startDate", "endDate"],
                range: 1,
                raw: false,
                dateNF: "dd/mm/yyyy"
            });

            const validationErrors = [];
            const validData = [];

            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const validAnnouncementTypes = (oAnnouncementTypeModel.getProperty("/types") || [])
                .map(t => t.text.toLowerCase());

            const oCategoryModel = this.getView().getModel("categoryModel");
            const validCategories = (oCategoryModel.getProperty("/category") || [])
                .map(c => c.text);
            const nameToIdMap = oCategoryModel.getProperty("/nameToIdMap") || {};

            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            jsonData.forEach((row, index) => {
                const rowNumber = index + 2;
                const title = row.title ? row.title.toString().trim() : "";
                const announcementType = row.announcementType ? row.announcementType.toString().trim() : "";
                const category = row.category ? row.category.toString().trim() : "";
                const description = row.description ? row.description.toString().trim() : "";
                const startDate = row.startDate ? row.startDate.toString().trim() : "";
                const endDate = row.endDate ? row.endDate.toString().trim() : "";

                let hasError = false;

                // Validate Title
                if (!title || title === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "Title",
                        error: "Title is required and cannot be empty"
                    });
                    hasError = true;
                }

                // Validate Announcement Type
                if (!announcementType || announcementType === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "Announcement Type",
                        error: "Announcement Type is required and cannot be empty"
                    });
                    hasError = true;
                } else {
                    const announcementTypes = announcementType.split(",").map(t => t.trim()).filter(Boolean);
                    const invalidTypes = [];

                    announcementTypes.forEach(type => {
                        if (!validAnnouncementTypes.includes(type.toLowerCase())) {
                            invalidTypes.push(type);
                        }
                    });

                    if (invalidTypes.length > 0) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "Announcement Type",
                            error: `Invalid announcement type(s): "${invalidTypes.join(", ")}". Valid types are: Global, Emergency, Process`
                        });
                        hasError = true;
                    }
                }

                // Validate Category
                if (!category || category === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "Category",
                        error: "Category is required and cannot be empty"
                    });
                    hasError = true;
                } else {
                    const categories = category.split(",").map(c => c.trim()).filter(Boolean);
                    const invalidCategories = [];

                    categories.forEach(cat => {
                        if (!nameToIdMap[cat]) {
                            invalidCategories.push(cat);
                        }
                    });

                    if (invalidCategories.length > 0) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "Category",
                            error: `Invalid category(s): "${invalidCategories.join(", ")}". Please use valid categories from the system`
                        });
                        hasError = true;
                    }
                }

                // Validate Description
                if (!description || description === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "Description",
                        error: "Description is required and cannot be empty"
                    });
                    hasError = true;
                }

                // Validate Start Date
                if (!startDate || startDate === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "Start Date",
                        error: "Start Date is required and cannot be empty"
                    });
                    hasError = true;
                } else {
                    const parsedStartDate = this._parseExcelDate(startDate);
                    if (!parsedStartDate) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "Start Date",
                            error: "Invalid date format. Use dd/MM/yyyy"
                        });
                        hasError = true;
                    } else if (parsedStartDate < oToday) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "Start Date",
                            error: "Start Date cannot be in the past"
                        });
                        hasError = true;
                    }
                }

                // Validate End Date
                if (!endDate || endDate === "") {
                    validationErrors.push({
                        row: rowNumber,
                        column: "End Date",
                        error: "End Date is required and cannot be empty"
                    });
                    hasError = true;
                } else {
                    const parsedEndDate = this._parseExcelDate(endDate);
                    if (!parsedEndDate) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "End Date",
                            error: "Invalid date format. Use dd/MM/yyyy"
                        });
                        hasError = true;
                    } else if (parsedEndDate < oToday) {
                        validationErrors.push({
                            row: rowNumber,
                            column: "End Date",
                            error: "End Date cannot be in the past"
                        });
                        hasError = true;
                    } else if (startDate) {
                        const parsedStartDate = this._parseExcelDate(startDate);
                        if (parsedStartDate && parsedEndDate < parsedStartDate) {
                            validationErrors.push({
                                row: rowNumber,
                                column: "End Date",
                                error: "End Date must be on or after Start Date"
                            });
                            hasError = true;
                        }
                    }
                }

                if (!hasError) {
                    validData.push({
                        title: title,
                        announcementType: announcementType,
                        category: category,
                        description: description,
                        startDate: this._formatDateToValue(this._parseExcelDate(startDate)),
                        endDate: this._formatDateToValue(this._parseExcelDate(endDate))
                    });
                }
            });

            return {
                validData: validData,
                errors: validationErrors,
                totalRows: jsonData.length
            };
        },

        _parseExcelDate: function (dateStr) {
            if (!dateStr) return null;

            // Try parsing dd/MM/yyyy format
            const parts = dateStr.split('/');
            if (parts.length === 3) {
                const day = parseInt(parts[0], 10);
                const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed
                const year = parseInt(parts[2], 10);

                if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                    const date = new Date(year, month, day);
                    if (date.getDate() === day && date.getMonth() === month && date.getFullYear() === year) {
                        return date;
                    }
                }
            }

            // Try parsing as ISO date
            const isoDate = new Date(dateStr);
            if (!isNaN(isoDate.getTime())) {
                return isoDate;
            }

            return null;
        },

        _showValidationErrorDialog: function (errors, validCount, totalCount) {
            const errorMessages = errors.map(err =>
                `Row ${err.row}, Column '${err.column}': ${err.error}`
            ).join("\n");

            const fullMessage = `Validation Failed!\n\n` +
                `Total Rows: ${totalCount}\n` +
                `Valid Rows: ${validCount}\n` +
                `Invalid Rows: ${errors.length}\n\n` +
                `Errors:\n${errorMessages}\n\n` +
                `Please correct the errors in your Excel file and upload again.`;

            MessageBox.error(fullMessage, {
                title: "Excel Validation Error",
                contentWidth: "500px",
                styleClass: "sapUiSizeCompact"
            });
        },

        _setUploadStatus: function (oModel, sStatus, sMessage, sType) {
            oModel.setProperty("/uploadStatus", sStatus);
            oModel.setProperty("/uploadMessage", sMessage);
            oModel.setProperty("/uploadMessageType", sType);
            oModel.setProperty("/showUploadMessage", true);
            oModel.setProperty("/templateUploaded", sStatus === "Uploaded" ? "Yes" : "No");
            oModel.setProperty("/bulkUploadStepValidated", sStatus === "Uploaded");
        },

        _handleSuccessfulUpload: function (oWizardModel, bulkData) {
            oWizardModel.setProperty("/bulkData", bulkData);
            this._setUploadStatus(
                oWizardModel,
                "Uploaded",
                `Excel file uploaded successfully! ${bulkData.length} record(s) loaded. Please review and edit in the next step.`,
                "Success"
            );

            MessageToast.show(`Excel file uploaded successfully! ${bulkData.length} record(s) loaded.`);

            const oBulkWizard = this.byId("bulkWizard");
            if (oBulkWizard) {
                setTimeout(() => {
                    oBulkWizard.nextStep();

                    // FIXED: Button visibility will be handled by onBulkReviewStepActivate
                    // Don't set button visibility here
                }, 1000);
            }
        },

        _handleUploadError: function (oWizardModel, error) {
            MessageBox.error("Error reading Excel file: " + error.message);
            oWizardModel.setProperty("/uploadStatus", "Failed");
            oWizardModel.setProperty("/uploadMessage", "Error reading Excel file.");
            oWizardModel.setProperty("/uploadMessageType", "Error");
        },

        /* ========================================
         * EDIT FUNCTIONALITY
         * ======================================== */


        onEditPress: function (oEvent) {
            const oButton = oEvent.getSource();
            const oListItem = oButton.getParent().getParent();
            const oBindingContext = oListItem.getBindingContext();

            if (!oBindingContext) {
                MessageBox.error("Unable to get announcement data. Please refresh and try again.");
                return;
            }

            const sPath = oBindingContext.getPath();
            const oModel = this.getOwnerComponent().getModel("announcementModel");

            if (!oModel) {
                MessageBox.error("Model not found. Please refresh and try again.");
                return;
            }

            // Show busy indicator
            const oBusy = new sap.m.BusyDialog({ text: "Loading announcement..." });
            oBusy.open();

            // Read the announcement with expanded toTypes
            oModel.read(sPath, {
                urlParameters: {
                    "$expand": "toTypes"
                },
                success: (oData) => {
                    oBusy.close();

                    console.log("Edit Data:", oData);
                    console.log("toTypes:", oData.toTypes);

                    // Extract typeIds with improved logic
                    const aTypeIds = this._extractTypeIds(oData.toTypes);

                    console.log("Extracted typeIds:", aTypeIds);

                    const oEditData = {
                        id: oData.announcementId,
                        title: oData.title,
                        announcementType: oData.announcementType,
                        description: oData.description,
                        createdBy: oData.createdBy,
                        createdOn: oData.createdAt,
                        modifiedBy: oData.modifiedBy,
                        modifiedOn: oData.modifiedAt,
                        publishedAt: oData.publishedAt,
                        publishedBy: oData.publishedBy,
                        announcementStatus: oData.announcementStatus,
                        publishStatus: this._determinePublishStatus(oData),
                        startAnnouncement: oData.startAnnouncement,
                        endAnnouncement: oData.endAnnouncement,
                        isActive: oData.isActive,
                        typeId: aTypeIds.length > 0 ? aTypeIds : []
                    };

                    console.log("Final Edit Data:", oEditData);

                    this._editContext = {
                        data: oEditData,
                        path: sPath,
                        isEdit: true
                    };

                    this._openWizardForEdit(oEditData);
                },
                error: (oError) => {
                    oBusy.close();
                    console.error("Failed to read announcement:", oError);
                    MessageBox.error("Failed to load announcement data. Please try again.");
                }
            });
        },

        _extractTypeIds: function (toTypes) {
            if (!toTypes) {
                return [];
            }

            // Handle OData V2 structure with __list, results, or direct array
            let aResults = toTypes.__list || toTypes.results || toTypes;

            if (!Array.isArray(aResults)) {
                return [];
            }

            return aResults
                .map(item => {
                    // Try multiple possible paths for typeId
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
        _determinePublishStatus: function (oData) {
            if (!oData.startAnnouncement || !oData.endAnnouncement) {
                return "";
            }

            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);
            const oStartDate = new Date(oData.startAnnouncement);
            oStartDate.setHours(0, 0, 0, 0);

            if (oStartDate.getTime() === oToday.getTime()) {
                return "PUBLISH_NOW";
            } else if (oStartDate > oToday) {
                return "PUBLISH_LATER";
            }

            return "";
        },

        _openWizardForEdit: function (oEditData) {
            this._loadWizardDialog()
                .then(() => {
                    // Initialize wizard model with edit data
                    this._initWizardModelForEdit(oEditData);

                    // Reset wizards to clean state
                    this._resetWizards();

                    // Open dialog first
                    this._oWizardDialog.open();

                    // Wait for dialog to fully render before setting up wizard
                    // Increased timeout to ensure DOM is ready
                    setTimeout(() => {
                        this._setupSingleWizardForEdit();
                    }, 300); // Increased from 150ms to 300ms
                })
                .catch(error => {
                    MessageBox.error("Failed to open edit dialog: " + error.message);
                });
        },

        _ensureWizardContentVisible: function () {
            // Force revalidation of all input controls
            const aInputIds = [
                "singleTitleInput",
                "singleAnnouncementTypeMultiComboBox",
                "singleCategoryMultiComboBox"
            ];

            aInputIds.forEach(sInputId => {
                const oInput = this.byId(sInputId);
                if (oInput) {
                    oInput.invalidate();
                }
            });

            // Force revalidation of the rich text container
            const oContainer = this.byId("richTextContainer");
            if (oContainer) {
                oContainer.invalidate();
            }
        },

        _initializeEditWizard: function (oEditData) {
            this._initWizardModelForEdit(oEditData);
            this._resetWizards();

            // Open dialog first
            this._oWizardDialog.open();

            // Setup wizard after dialog is visible
            setTimeout(() => {
                this._setupSingleWizardForEdit();
            }, 100);
        },

        _setupSingleWizardForEdit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const oWizard = this.byId("singleWizard");
            const oSelectionStep = this.byId("singleSelectionStep");
            const oCreateStep = this.byId("singleCreateStep");

            if (!oWizard || !oCreateStep || !oSelectionStep) {
                console.error("Wizard or steps not found");
                return;
            }

            // Hide selection step (Edit mode)
            oSelectionStep.setVisible(false);
            oSelectionStep.setValidated(true);

            // Validate create step
            oCreateStep.setValidated(true);

            // Set current step
            oWizard.setCurrentStep(oCreateStep);

            // Set footer buttons for edit mode
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: false,
                showSubmitButton: true,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false,
                showResetButton: false  // Will be shown when user makes changes
            });

            // Force wizard layout update
            setTimeout(() => {
                oWizard.invalidate();

                // Initialize RTE after wizard is fully rendered
                setTimeout(() => {
                    const sDescription = oWizardModel.getProperty("/description") || "";
                    console.log("Initializing RTE with description:", sDescription);

                    this._initRichTextEditor("richTextContainer");

                    // Verify RTE was created and has the right value
                    if (this._oRichTextEditor) {
                        console.log("RTE initialized, value:", this._oRichTextEditor.getValue());
                    }

                    // FIX: After RTE is initialized, check if Reset button should be visible
                    // This ensures Reset button appears if original data differs from defaults
                    this._handleResetButtonVisibility(oWizardModel);
                }, 200);
            }, 200);
        },


        /* ========================================
         * SUBMIT HANDLERS
         * ======================================== */

        // onPublishPress: function () {
        //     const oWizardModel = this.getView().getModel("wizardModel");
        //     const sCurrentFlow = oWizardModel.getProperty("/currentFlow");
        //     const bIsEditMode = oWizardModel.getProperty("/isEditMode");

        //     // Validate publish options for BOTH flows
        //     const bValidPublish = this._validatePublishOptions();
        //     if (!bValidPublish) {
        //         return;
        //     }

        //     // Show confirmation dialog before publishing
        //     const sConfirmMessage = bIsEditMode
        //         ? "Are you sure you want to publish these changes?"
        //         : (sCurrentFlow === "BULK"
        //             ? "Are you sure you want to publish these announcements?"
        //             : "Are you sure you want to publish this announcement?");

        //     MessageBox.confirm(sConfirmMessage, {
        //         title: "Confirm Publish",
        //         actions: [MessageBox.Action.YES, MessageBox.Action.NO],
        //         emphasizedAction: MessageBox.Action.YES,
        //         onClose: (oAction) => {
        //             if (oAction === MessageBox.Action.YES) {
        //                 // Call submit handlers
        //                 if (sCurrentFlow === "SINGLE") {
        //                     bIsEditMode ? this._handleEditSubmit() : this._handleSingleSubmit();
        //                 } else if (sCurrentFlow === "BULK") {
        //                     this._handleBulkSubmit();
        //                 }
        //             }
        //         }
        //     });
        // },

        onDraftPress: function () {
            // if (!this._validateSingleCreateStep()) {
            //     MessageToast.show("Please complete all required fields with valid information");
            //     return;
            // }

            // MessageBox.confirm(
            //     "Are you sure you want to save this as draft?",
            //     {
            //         title: "Confirm Draft",
            //         actions: [MessageBox.Action.YES, MessageBox.Action.NO],
            //         emphasizedAction: MessageBox.Action.YES,
            //         onClose: (oAction) => {
            //             if (oAction === MessageBox.Action.YES) {
            //                 this._handleDraftSubmit();
            //             }
            //         }
            //     }
            // );

            this._handleDraftSubmit();
        },

        // _handleDraftSubmit: function () {
        //     const oWizardModel = this.getView().getModel("wizardModel");
        //     const sTitle = (oWizardModel.getProperty("/title") || "").trim();
        //     const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
        //     const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
        //     const aTypeIds = oWizardModel.getProperty("/category") || [];

        //     const sDescription = this._stripHtmlTags(sDescriptionHTML);
        //     const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
        //     const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];
        //     const aSelectedTypes = aAnnouncementTypeKeys
        //         .map(key => aTypeList.find(t => t.key === key)?.text)
        //         .filter(Boolean);
        //     const sAnnouncementType = aSelectedTypes.join(",");

        //     const sPublishOption = oWizardModel.getProperty("/publishOption");
        //     let startAnnouncement = null;
        //     let endAnnouncement = null;

        //     if (sPublishOption === "PUBLISH") {
        //         const sEndDate = oWizardModel.getProperty("/publishEndDate");
        //         if (sEndDate) {
        //             startAnnouncement = new Date().toISOString();
        //             const oEndDate = new Date(sEndDate);
        //             oEndDate.setDate(oEndDate.getDate() + 1);
        //             endAnnouncement = oEndDate.toISOString();
        //         }
        //     }

        //     const oPayload = {
        //         data: [{
        //             title: sTitle,
        //             description: sDescription,
        //             announcementType: sAnnouncementType,
        //             announcementStatus: "DRAFT",
        //             startAnnouncement: startAnnouncement,
        //             endAnnouncement: endAnnouncement,
        //             publishedBy: null,
        //             toTypes: aTypeIds.map(typeId => ({
        //                 type: { typeId: typeId }
        //             }))
        //         }]
        //     };

        //     const oBusy = new sap.m.BusyDialog({ text: "Saving as draft..." });
        //     oBusy.open();

        //     const oAnnouncementModel = this.getOwnerComponent().getModel("announcementModel");

        //     oAnnouncementModel.callFunction("/bulkCreateAnnouncements", {
        //         method: "POST",
        //         urlParameters: oPayload,
        //         success: (oData) => {
        //             oBusy.close();
        //             this._oWizardDialog.close();
        //             sap.m.MessageToast.show("Announcement saved as draft successfully!");
        //             this.refreshSmartTable();
        //         },
        //         error: (oError) => {
        //             oBusy.close();
        //             console.error("Save as draft failed:", oError);
        //             sap.m.MessageBox.error("Failed to save as draft.");
        //         }
        //     });
        // },
        _validatePublishOptions: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");

            // BULK flow doesn't need publish option validation (dates are in the table)
            if (sCurrentFlow === "BULK") {
                return true; // Validation happens in _validateBulkTableData
            }

            // SINGLE flow validation
            const sOption = oWizardModel.getProperty("/publishOption");
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            const sEndDate = oWizardModel.getProperty("/publishEndDate");
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            // Clear previous error states
            oWizardModel.setProperty("/publishStartDateValueState", "None");
            oWizardModel.setProperty("/publishStartDateValueStateText", "");
            oWizardModel.setProperty("/publishEndDateValueState", "None");
            oWizardModel.setProperty("/publishEndDateValueStateText", "");

            if (!sOption) {
                MessageBox.error("Please select whether you want to Publish now or Publish later.");
                return false;
            }

            if (sOption === "UNPUBLISH") {
                // Validate Start Date FIRST for Publish Later
                if (!sStartDate) {
                    oWizardModel.setProperty("/publishStartDateValueState", "Error");
                    oWizardModel.setProperty("/publishStartDateValueStateText", "Start Date is required for Publish Later");
                    MessageBox.error("Please select a Start Date for Publish Later.");
                    return false;
                }

                // Validate Start Date is not in the past
                const oStartDate = new Date(sStartDate);
                oStartDate.setHours(0, 0, 0, 0);
                if (oStartDate < oToday) {
                    oWizardModel.setProperty("/publishStartDateValueState", "Error");
                    oWizardModel.setProperty("/publishStartDateValueStateText", "Start Date cannot be in the past");
                    MessageBox.error("Start Date cannot be in the past.");
                    return false;
                }
            }

            // Validate End Date (required for both options)
            if (!sEndDate) {
                oWizardModel.setProperty("/publishEndDateValueState", "Error");
                oWizardModel.setProperty("/publishEndDateValueStateText", "End Date is required");
                MessageBox.error("Please select an End Date.");
                return false;
            }

            // Validate End Date is not in the past
            const oEndDate = new Date(sEndDate);
            oEndDate.setHours(0, 0, 0, 0);
            if (oEndDate < oToday) {
                oWizardModel.setProperty("/publishEndDateValueState", "Error");
                oWizardModel.setProperty("/publishEndDateValueStateText", "End Date cannot be in the past");
                MessageBox.error("End Date cannot be in the past.");
                return false;
            }

            // **CHANGED: Validate End Date is AFTER Start Date (not equal)**
            if (sOption === "UNPUBLISH") {
                const oStartDate = new Date(sStartDate);
                oStartDate.setHours(0, 0, 0, 0);
                if (oEndDate <= oStartDate) { // **CHANGED: from < to <=**
                    oWizardModel.setProperty("/publishEndDateValueState", "Error");
                    oWizardModel.setProperty("/publishEndDateValueStateText", "End Date must be after Start Date");
                    MessageBox.error("End Date must be after Start Date.");
                    return false;
                }
            } else if (sOption === "PUBLISH") {
                // For Publish Now, end date must be after today
                if (oEndDate <= oToday) { // **CHANGED: from < to <=**
                    oWizardModel.setProperty("/publishEndDateValueState", "Error");
                    oWizardModel.setProperty("/publishEndDateValueStateText", "End Date must be after today");
                    MessageBox.error("End Date must be after today.");
                    return false;
                }
            }

            return true;
        },
        // onCancelPress: function () {
        //     const oWizard = this.byId("singleWizard");
        //     if (oWizard) {
        //         oWizard.removeStyleClass("hideFirstWizardStep");
        //     }

        //     // CLEANUP BOTH RICHTEXTEDITORS
        //     if (this._oRichTextEditor) {
        //         this._oRichTextEditor.destroy();
        //         this._oRichTextEditor = null;
        //     }

        //     // ADD THIS:
        //     if (this._oReviewRichTextEditor) {
        //         this._oReviewRichTextEditor.destroy();
        //         this._oReviewRichTextEditor = null;
        //     }

        //     this._editContext = null;
        //     this._oWizardDialog.close();
        // },

        _stripHtmlTags: function (sHtml) {
            if (!sHtml) return "";

            // Create a temporary div element to parse HTML
            const tmp = document.createElement("DIV");
            tmp.innerHTML = sHtml;

            // 1. REMOVE SCRIPT AND STYLE TAGS completely (do this first)
            const scriptsAndStyles = tmp.querySelectorAll('script, style');
            scriptsAndStyles.forEach(el => el.remove());

            // 2. HANDLE IMAGES - Replace with alt text or placeholder
            const images = tmp.querySelectorAll('img');
            images.forEach(img => {
                const altText = img.getAttribute('alt') || '[Image]';
                const textNode = document.createTextNode(altText + ' ');
                img.parentNode.replaceChild(textNode, img);
            });

            // 3. HANDLE LINKS - Keep link text but remove href
            const links = tmp.querySelectorAll('a');
            links.forEach(link => {
                const textNode = document.createTextNode(link.textContent || '[Link]');
                link.parentNode.replaceChild(textNode, link);
            });

            // 4. HANDLE LISTS - Convert to comma-separated format (NO bullets)
            const lists = tmp.querySelectorAll('ul, ol');
            lists.forEach(list => {
                const listItems = list.querySelectorAll('li');
                const items = [];
                listItems.forEach(li => {
                    const itemText = li.textContent.trim();
                    if (itemText) {
                        items.push(itemText);
                    }
                });
                const textNode = document.createTextNode(items.join(', ') + ' ');
                list.parentNode.replaceChild(textNode, list);
            });

            // 5. HANDLE TABLES - Extract cell content with separators
            const tables = tmp.querySelectorAll('table');
            tables.forEach(table => {
                const rows = table.querySelectorAll('tr');
                const rowTexts = [];
                rows.forEach(row => {
                    const cells = row.querySelectorAll('td, th');
                    const cellTexts = [];
                    cells.forEach(cell => {
                        const cellText = cell.textContent.trim();
                        if (cellText) {
                            cellTexts.push(cellText);
                        }
                    });
                    if (cellTexts.length > 0) {
                        rowTexts.push(cellTexts.join(' | '));
                    }
                });
                const textNode = document.createTextNode(rowTexts.join('; ') + ' ');
                table.parentNode.replaceChild(textNode, table);
            });

            // 6. HANDLE BLOCK ELEMENTS - Add space after each block
            const blockElements = tmp.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, blockquote, section, article, header, footer, nav, aside');
            blockElements.forEach(el => {
                const text = el.textContent || '';
                el.textContent = text + ' ';
            });

            // 7. HANDLE LINE BREAKS - Replace with spaces
            tmp.innerHTML = tmp.innerHTML.replace(/<br\s*\/?>/gi, ' ');

            // 8. Get final text content
            let text = tmp.textContent || tmp.innerText || "";

            // 9. COMPREHENSIVE CLEAN UP:
            text = text
                // Remove all newlines, returns, tabs
                .replace(/[\n\r\t]/g, ' ')
                // Remove all types of bullet characters (•, ◦, ▪, ▫, ‣, ⁃, ⁌, ⁍, ●, ○, ■, □, ▶, ►)
                .replace(/[•◦▪▫‣⁃⁌⁍●○■□▶►]/g, '')
                // Remove common list markers (-, *, +, >) at the start of text segments
                .replace(/(?:^|\s)[-*+>]\s+/g, ' ')
                // Remove numbered list markers (1., 2., a., b., i., ii., etc.)
                .replace(/(?:^|\s)\d+\.\s+/g, ' ')
                .replace(/(?:^|\s)[a-z]\.\s+/gi, ' ')
                .replace(/(?:^|\s)[ivx]+\.\s+/gi, ' ')
                // Remove HTML entities that might remain
                .replace(/&nbsp;/g, ' ')
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&#39;/g, "'")
                // Remove zero-width spaces and other invisible characters
                .replace(/[\u200B-\u200D\uFEFF]/g, '')
                // Remove control characters (except normal space)
                .replace(/[\x00-\x1F\x7F-\x9F]/g, '')
                // Replace multiple spaces with single space
                .replace(/\s{2,}/g, ' ')
                // Remove leading/trailing spaces
                .trim();

            // 10. Additional cleanup for edge cases
            // Remove repeated punctuation
            text = text.replace(/([.,;:!?])\1+/g, '$1');

            // Ensure proper spacing after punctuation
            text = text.replace(/([.,;:!?])(\S)/g, '$1 $2');

            // Remove spaces before punctuation
            text = text.replace(/\s+([.,;:!?])/g, '$1');

            // 11. Limit length if needed
            const MAX_LENGTH = 5000;
            if (text.length > MAX_LENGTH) {
                text = text.substring(0, MAX_LENGTH).trim() + '...';
            }

            return text;
        },

        // _handleSingleSubmit: function () {
        //     const oView = this.getView();
        //     const oWizardModel = oView.getModel("wizardModel");

        //     const sTitle = (oWizardModel.getProperty("/title") || "").trim();
        //     const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
        //     const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
        //     const aTypeIds = oWizardModel.getProperty("/category") || [];
        //     const sPublishOption = oWizardModel.getProperty("/publishOption");
        //     const sStartDate = oWizardModel.getProperty("/publishStartDate");
        //     const sEndDate = oWizardModel.getProperty("/publishEndDate");

        //     if (!sTitle || !sDescriptionHTML || !aAnnouncementTypeKeys.length || !aTypeIds.length) {
        //         sap.m.MessageToast.show("Please fill in all required fields");
        //         return;
        //     }

        //     const sDescription = this._stripHtmlTags(sDescriptionHTML);

        //     const oAnnouncementTypeModel = oView.getModel("announcementTypeModel");
        //     const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];

        //     const sAnnouncementType = aAnnouncementTypeKeys
        //         .map(key => aTypeList.find(t => t.key === key)?.text)
        //         .filter(Boolean)
        //         .join(",");

        //     this.getCurrentUserEmail().then((sUserEmail) => {

        //         let announcementStatus, startAnnouncement, endAnnouncement;
        //         const nowISO = new Date().toISOString();

        //         if (sPublishOption === "PUBLISH") {
        //             announcementStatus = "PUBLISHED";
        //             startAnnouncement = nowISO;
        //         } else {
        //             announcementStatus = "TO_BE_PUBLISHED";
        //             startAnnouncement = new Date(sStartDate).toISOString();
        //         }

        //         const oEndDate = new Date(sEndDate);
        //         oEndDate.setDate(oEndDate.getDate() + 1);
        //         endAnnouncement = oEndDate.toISOString();

        //         const oPayload = {
        //             data: [{
        //                 title: sTitle,
        //                 description: sDescription,
        //                 announcementType: sAnnouncementType,
        //                 announcementStatus: announcementStatus,
        //                 startAnnouncement: startAnnouncement,
        //                 endAnnouncement: endAnnouncement,
        //                 publishedBy: sUserEmail,
        //                 toTypes: aTypeIds.map(typeId => ({
        //                     type: { typeId }
        //                 }))
        //             }]
        //         };

        //         const oBusy = new sap.m.BusyDialog({ text: "Creating announcement..." });
        //         oBusy.open();

        //         const oModel = this.getOwnerComponent().getModel("announcementModel");

        //         // Intercept before request is sent to modify it
        //         const fnBeforeRequestSent = function (oEvent) {
        //             const oRequest = oEvent.getParameter("request");
        //             const sUrl = oRequest.requestUri || oRequest.url;

        //             if (sUrl && sUrl.includes("bulkCreateAnnouncements")) {
        //                 // Remove query parameters from URL
        //                 if (oRequest.requestUri) {
        //                     oRequest.requestUri = oRequest.requestUri.split('?')[0];
        //                 }
        //                 if (oRequest.url) {
        //                     oRequest.url = oRequest.url.split('?')[0];
        //                 }

        //                 // Set the payload as request body
        //                 oRequest.data = JSON.stringify(oPayload);

        //                 // Ensure correct headers
        //                 oRequest.headers = oRequest.headers || {};
        //                 oRequest.headers["Content-Type"] = "application/json";
        //                 oRequest.headers["Accept"] = "application/json";
        //             }
        //         };

        //         // Attach the interceptor
        //         oModel.attachRequestSent(fnBeforeRequestSent);

        //         // Call the function without urlParameters
        //         oModel.callFunction("/bulkCreateAnnouncements", {
        //             method: "POST",
        //             success: (oData) => {
        //                 oBusy.close();
        //                 oModel.detachRequestSent(fnBeforeRequestSent);
        //                 this._oWizardDialog.close();

        //                 const oResult = oData?.bulkCreateAnnouncements;

        //                 sap.m.MessageToast.show(
        //                     oResult?.message || "Announcement created successfully"
        //                 );

        //                 this.refreshSmartTable();
        //             },
        //             error: (oError) => {
        //                 oBusy.close();
        //                 oModel.detachRequestSent(fnBeforeRequestSent);

        //                 let sMessage = "Failed to create announcement";
        //                 try {
        //                     const oErr = JSON.parse(oError.responseText);
        //                     sMessage = oErr?.error?.message?.value || sMessage;
        //                 } catch (e) { /* ignore */ }

        //                 sap.m.MessageBox.error(sMessage);
        //             }
        //         });

        //     }).catch((err) => {
        //         sap.m.MessageBox.error("Failed to get current user: " + err.message);
        //     });
        // },

        _handleEditSubmit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aCategories = oWizardModel.getProperty("/category") || [];
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const sEditId = oWizardModel.getProperty("/editId");

            const sPublishOption = oWizardModel.getProperty("/publishOption");
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            const sEndDate = oWizardModel.getProperty("/publishEndDate");

            if (!sTitle || aAnnouncementTypeKeys.length === 0 || aCategories.length === 0 || !sDescriptionHTML) {
                sap.m.MessageToast.show("Please fill in all required fields");
                return;
            }

            const sDescription = this._stripHtmlTags(sDescriptionHTML);
            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];
            const aSelectedTypes = aAnnouncementTypeKeys
                .map(key => aTypeList.find(t => t.key === key)?.text)
                .filter(Boolean);
            const sAnnouncementType = aSelectedTypes.join(",");

            const oBusy = new sap.m.BusyDialog({ text: "Updating announcement..." });
            oBusy.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement;
                    const currentDateTime = new Date().toISOString();

                    if (sPublishOption === "PUBLISH") {
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    } else {
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sStartDate).toISOString();
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    }

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
                        toTypes: aCategories.map(typeId => ({
                            type: { typeId: typeId }
                        }))
                    };

                    // Fetch CSRF token first
                    this._getCSRFToken()
                        .then((csrfToken) => {
                            $.ajax({
                                url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${sEditId}')`,
                                method: "PATCH",
                                contentType: "application/json",
                                dataType: "json",
                                headers: {
                                    "X-CSRF-Token": csrfToken  //Include CSRF token
                                },
                                data: JSON.stringify(oPayload),
                                success: (oResponse) => {
                                    oBusy.close();

                                    const oWizard = this.byId("singleWizard");
                                    if (oWizard) {
                                        oWizard.removeStyleClass("hideFirstWizardStep");
                                    }

                                    this._editContext = null;
                                    this._oWizardDialog.close();

                                    const sMessage = announcementStatus === "PUBLISHED"
                                        ? `Announcement '${sTitle}' updated and published successfully!`
                                        : `Announcement '${sTitle}' updated and scheduled for publication!`;

                                    sap.m.MessageToast.show(sMessage);

                                    setTimeout(() => {
                                        this.refreshSmartTable();
                                    }, 500);
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Update announcement failed:", status, err);
                                    console.error("Response:", xhr.responseText);
                                    let sErrorMessage = "Failed to update announcement. Please try again.";
                                    if (xhr.responseJSON?.error?.message) {
                                        sErrorMessage = xhr.responseJSON.error.message;
                                    }
                                    sap.m.MessageBox.error(sErrorMessage);
                                }
                            });
                        })
                        .catch((err) => {
                            oBusy.close();
                            console.error("CSRF token fetch failed:", err);
                            sap.m.MessageBox.error("Failed to initialize request. Please try again.");
                        });
                })
                .catch((error) => {
                    oBusy.close();
                    sap.m.MessageBox.error("Failed to get current user: " + error.message);
                });
        },

        _handleBulkSubmit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const oCategoryModel = this.getView().getModel("categoryModel");
            const bulkData = oWizardModel.getProperty("/bulkData") || [];

            if (bulkData.length === 0) {
                sap.m.MessageBox.error("No data to submit. Please upload a valid Excel file.");
                return;
            }

            const validationResult = this._validateBulkTableData(bulkData);
            if (!validationResult.isValid) {
                this._showBulkValidationErrorDialog(validationResult.errors);
                return;
            }

            const nameToIdMap = oCategoryModel.getProperty("/nameToIdMap") || {};
            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    const aPayloads = [];
                    const aInvalidItems = [];

                    bulkData.forEach((item, index) => {
                        const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
                        const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];
                        const sAnnouncementType = (item.announcementType || "")
                            .split(",")
                            .map(name => name.trim())
                            .map(name => {
                                const match = aTypeList.find(t => t.text.toLowerCase() === name.toLowerCase());
                                return match ? match.text : name;
                            })
                            .join(", ");

                        const aCategoryNames = (item.category || "")
                            .split(",")
                            .map(name => name.trim())
                            .filter(Boolean);

                        const aTypeIds = aCategoryNames
                            .map(name => nameToIdMap[name])
                            .filter(Boolean);

                        if (aTypeIds.length === 0) {
                            aInvalidItems.push({
                                index: index + 1,
                                title: item.title,
                                reason: `Invalid category: ${item.category}`
                            });
                            return;
                        }

                        if (!sAnnouncementType) {
                            aInvalidItems.push({
                                index: index + 1,
                                title: item.title,
                                reason: "Announcement Type is required"
                            });
                            return;
                        }

                        const oStartDate = new Date(item.startDate);
                        oStartDate.setHours(0, 0, 0, 0);

                        let announcementStatus, startAnnouncement, endAnnouncement;
                        const currentDateTime = new Date().toISOString();

                        if (oStartDate.getTime() === oToday.getTime()) {
                            announcementStatus = "PUBLISHED";
                            startAnnouncement = currentDateTime;
                        } else {
                            announcementStatus = "TO_BE_PUBLISHED";
                            startAnnouncement = new Date(item.startDate).toISOString();
                        }

                        const oEndDate = new Date(item.endDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();

                        aPayloads.push({
                            title: item.title.trim(),
                            description: item.description.trim(),
                            announcementType: sAnnouncementType,
                            announcementStatus: announcementStatus,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishedBy: sUserEmail,
                            toTypes: aTypeIds.map(typeId => ({
                                type: { typeId: typeId }
                            }))
                        });
                    });

                    if (aInvalidItems.length > 0) {
                        const sErrorMsg = aInvalidItems
                            .map((item) => `Row ${item.index}: ${item.title} - ${item.reason}`)
                            .join("\n");

                        sap.m.MessageBox.error(`Cannot submit bulk announcements. Invalid data found:\n\n${sErrorMsg}`);
                        return;
                    }

                    const oBusy = new sap.m.BusyDialog({
                        text: `Publishing ${aPayloads.length} announcement(s)...`
                    });
                    oBusy.open();

                    const oPayload = {
                        data: aPayloads
                    };

                    // Fetch CSRF token first
                    this._getCSRFToken()
                        .then((csrfToken) => {
                            $.ajax({
                                url: "/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/bulkCreateAnnouncements",
                                method: "POST",
                                contentType: "application/json",
                                dataType: "json",
                                headers: {
                                    "X-CSRF-Token": csrfToken  // Include CSRF token
                                },
                                data: JSON.stringify(oPayload),
                                success: (oResponse) => {
                                    oBusy.close();

                                    const publishedCount = aPayloads.filter(p => p.announcementStatus === "PUBLISHED").length;
                                    const scheduledCount = aPayloads.filter(p => p.announcementStatus === "TO_BE_PUBLISHED").length;

                                    let sMessage = `Bulk upload successful! `;
                                    if (publishedCount > 0) {
                                        sMessage += `${publishedCount} announcement(s) published`;
                                    }
                                    if (scheduledCount > 0) {
                                        if (publishedCount > 0) sMessage += `, `;
                                        sMessage += `${scheduledCount} announcement(s) scheduled`;
                                    }

                                    sap.m.MessageToast.show(sMessage);
                                    this._oWizardDialog.close();

                                    // Add delay before refresh
                                    setTimeout(() => {
                                        this.refreshSmartTable();
                                    }, 500);
                                },
                                error: (xhr, status, err) => {
                                    oBusy.close();
                                    console.error("Bulk upload failed:", status, err);
                                    console.error("Response:", xhr.responseText);
                                    let sErrorMessage = "Bulk upload failed. Please check the data and try again.";
                                    if (xhr.responseJSON?.error?.message) {
                                        sErrorMessage = xhr.responseJSON.error.message;
                                    }
                                    sap.m.MessageBox.error(sErrorMessage);
                                }
                            });
                        })
                        .catch((err) => {
                            oBusy.close();
                            console.error("CSRF token fetch failed:", err);
                            sap.m.MessageBox.error("Failed to initialize request. Please try again.");
                        });
                })
                .catch((error) => {
                    sap.m.MessageBox.error("Failed to get current user: " + error.message);
                });
        },


        _validateBulkTableData: function (bulkData) {
            const errors = [];
            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const validAnnouncementTypes = (oAnnouncementTypeModel.getProperty("/types") || [])
                .map(t => t.text.toLowerCase());

            const oCategoryModel = this.getView().getModel("categoryModel");
            const nameToIdMap = oCategoryModel.getProperty("/nameToIdMap") || {};

            const oToday = new Date();
            oToday.setHours(0, 0, 0, 0);

            bulkData.forEach((row, index) => {
                const rowNumber = index + 1;

                // Validate Title
                if (!row.title || row.title.trim() === "") {
                    errors.push({ row: rowNumber, column: "Title", error: "Title is required" });
                }

                // Validate Announcement Type
                if (!row.announcementType || row.announcementType.trim() === "") {
                    errors.push({ row: rowNumber, column: "Announcement Type", error: "Announcement Type is required" });
                } else {
                    const types = row.announcementType.split(",").map(t => t.trim());
                    const invalidTypes = types.filter(t => !validAnnouncementTypes.includes(t.toLowerCase()));
                    if (invalidTypes.length > 0) {
                        errors.push({
                            row: rowNumber,
                            column: "Announcement Type",
                            error: `Invalid type(s): ${invalidTypes.join(", ")}`
                        });
                    }
                }

                // Validate Category
                if (!row.category || row.category.trim() === "") {
                    errors.push({ row: rowNumber, column: "Category", error: "Category is required" });
                } else {
                    const categories = row.category.split(",").map(c => c.trim());
                    const invalidCategories = categories.filter(c => !nameToIdMap[c]);
                    if (invalidCategories.length > 0) {
                        errors.push({
                            row: rowNumber,
                            column: "Category",
                            error: `Invalid category(s): ${invalidCategories.join(", ")}`
                        });
                    }
                }

                // Validate Description
                if (!row.description || row.description.trim() === "") {
                    errors.push({ row: rowNumber, column: "Description", error: "Description is required" });
                }

                // Validate Start Date
                if (!row.startDate || row.startDate.trim() === "") {
                    errors.push({ row: rowNumber, column: "Start Date", error: "Start Date is required" });
                } else {
                    const oStartDate = new Date(row.startDate);
                    oStartDate.setHours(0, 0, 0, 0);
                    if (isNaN(oStartDate.getTime())) {
                        errors.push({ row: rowNumber, column: "Start Date", error: "Invalid date format" });
                    } else if (oStartDate < oToday) {
                        errors.push({ row: rowNumber, column: "Start Date", error: "Start Date cannot be in the past" });
                    }
                }

                // Validate End Date
                if (!row.endDate || row.endDate.trim() === "") {
                    errors.push({ row: rowNumber, column: "End Date", error: "End Date is required" });
                } else {
                    const oEndDate = new Date(row.endDate);
                    oEndDate.setHours(0, 0, 0, 0);
                    if (isNaN(oEndDate.getTime())) {
                        errors.push({ row: rowNumber, column: "End Date", error: "Invalid date format" });
                    } else if (oEndDate < oToday) {
                        errors.push({ row: rowNumber, column: "End Date", error: "End Date cannot be in the past" });
                    } else if (row.startDate) {
                        const oStartDate = new Date(row.startDate);
                        oStartDate.setHours(0, 0, 0, 0);
                        if (!isNaN(oStartDate.getTime()) && oEndDate < oStartDate) {
                            errors.push({
                                row: rowNumber,
                                column: "End Date",
                                error: "End Date must be on or after Start Date"
                            });
                        }
                    }
                }
            });

            return {
                isValid: errors.length === 0,
                errors: errors
            };
        },

        _showBulkValidationErrorDialog: function (errors) {
            const errorMessages = errors.map(err =>
                `Row ${err.row}, ${err.column}: ${err.error}`
            ).join("\n");

            const fullMessage = `Validation Failed!\n\n` +
                `Total Errors: ${errors.length}\n\n` +
                `Errors:\n${errorMessages}\n\n` +
                `Please correct the errors in the table and try again.`;

            MessageBox.error(fullMessage, {
                title: "Bulk Data Validation Error",
                contentWidth: "500px",
                styleClass: "sapUiSizeCompact"
            });
        },

        /* ========================================
         * RESET HANDLERS
         * ======================================== */

        onResetPress: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");

            if (sCurrentFlow === "SINGLE") {
                MessageBox.confirm(
                    "Are you sure you want to reset the form? All entered data will be lost.",
                    {
                        title: "Confirm Reset",
                        actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                        emphasizedAction: MessageBox.Action.NO,
                        onClose: (oAction) => {
                            if (oAction === MessageBox.Action.YES) {
                                this._performSingleReset();
                            }
                        }
                    }
                );
            }
        },

        _performSingleReset: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");

            if (bIsEditMode) {
                // Edit Mode: Reset to original values
                const sOriginalTitle = oWizardModel.getProperty("/originalTitle") || "";
                const aOriginalAnnouncementType = oWizardModel.getProperty("/originalAnnouncementType") || [];
                const aOriginalCategory = oWizardModel.getProperty("/originalCategory") || [];
                const sOriginalDescription = oWizardModel.getProperty("/originalDescription") || "";

                // ✅ FIX: Also reset publish-related fields to original values
                const sOriginalPublishOption = oWizardModel.getProperty("/originalPublishOption") || "";
                const sOriginalPublishStartDate = oWizardModel.getProperty("/originalPublishStartDate") || "";
                const sOriginalPublishEndDate = oWizardModel.getProperty("/originalPublishEndDate") || "";

                // Reset basic fields
                oWizardModel.setProperty("/title", sOriginalTitle);
                oWizardModel.setProperty("/announcementType", [...aOriginalAnnouncementType]);
                oWizardModel.setProperty("/category", [...aOriginalCategory]);
                oWizardModel.setProperty("/description", sOriginalDescription);

                // ✅ FIX: Reset publish fields
                oWizardModel.setProperty("/publishOption", sOriginalPublishOption);
                oWizardModel.setProperty("/publishStartDate", sOriginalPublishStartDate);
                oWizardModel.setProperty("/publishEndDate", sOriginalPublishEndDate);

                // Reset date picker visibility and enabled states based on original publish option
                const bHasPublishData = !!(sOriginalPublishOption && sOriginalPublishEndDate);
                oWizardModel.setProperty("/showDatePickers", bHasPublishData);
                oWizardModel.setProperty("/startDateEnabled", sOriginalPublishOption === "UNPUBLISH");
                oWizardModel.setProperty("/endDateEnabled", bHasPublishData);

                // Update RichTextEditor value
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue(sOriginalDescription);
                }

                // Clear validation errors and validate step
                this._clearValidationErrors(oWizardModel);
                oWizardModel.setProperty("/singleCreateStepValidated", true);

                const oStep = this.byId("singleCreateStep");
                if (oStep) {
                    oStep.setValidated(true);
                }

                MessageToast.show("Reset to original values");
            } else {
                // Create Mode: Clear all fields
                oWizardModel.setProperty("/title", "");
                oWizardModel.setProperty("/announcementType", []);
                oWizardModel.setProperty("/category", []);
                oWizardModel.setProperty("/description", "");

                // ✅ FIX: Also clear publish fields
                oWizardModel.setProperty("/publishOption", "");
                oWizardModel.setProperty("/publishStartDate", "");
                oWizardModel.setProperty("/publishEndDate", "");
                oWizardModel.setProperty("/showDatePickers", false);
                oWizardModel.setProperty("/startDateEnabled", false);
                oWizardModel.setProperty("/endDateEnabled", false);

                // Update RichTextEditor value
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue("");
                }

                // Clear validation errors and invalidate step
                this._clearValidationErrors(oWizardModel);
                oWizardModel.setProperty("/singleCreateStepValidated", false);

                const oStep = this.byId("singleCreateStep");
                if (oStep) {
                    oStep.setValidated(false);
                }

                MessageToast.show("Form has been reset");
            }

            // Hide Reset button after reset
            oWizardModel.setProperty("/showResetButton", false);
        },

        /* ========================================
         * DELETE FUNCTIONALITY
         * ======================================== */

        onDeletePress: function (oEvent) {
            const oButton = oEvent.getSource();

            const oListItem = oButton.getParent().getParent();
            // const oBindingContext = oEvent.getSource().getBindingContext("announcementModel");
            const oBindingContext = oListItem.getBindingContext();

            if (!oBindingContext) {
                MessageBox.error("Unable to get announcement data. Please refresh and try again.");
                return;
            }

            const oData = oBindingContext.getObject();
            const sAnnouncementId = oData.announcementId;
            const sTitle = oData.title;

            if (!sAnnouncementId) {
                MessageBox.error("Unable to delete: Announcement ID not found.");
                return;
            }

            MessageBox.confirm(`Are you sure you want to delete '${sTitle}'?`, {
                title: "Confirm Delete",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.NO,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        this._deleteItem(sAnnouncementId, sTitle);
                    }
                }
            });
        },

        // Replace the _deleteItem method in your controller with this updated version:

        _deleteItem: function (sAnnouncementId, sTitle) {
            const oBusy = new sap.m.BusyDialog({ text: "Deleting announcement..." });
            oBusy.open();

            // Fetch CSRF token first
            this._getCSRFToken()
                .then((csrfToken) => {
                    $.ajax({
                        url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${sAnnouncementId}')`,
                        method: "DELETE",
                        contentType: "application/json",
                        dataType: "json",
                        headers: {
                            "X-CSRF-Token": csrfToken
                        },
                        success: (oResponse) => {
                            oBusy.close();

                            // Show centered MessageToast with longer duration, custom width and padding
                            sap.m.MessageToast.show(`Announcement '${sTitle}' deleted successfully!`, {
                                duration: 4000,                    // Duration in milliseconds (4 seconds)
                                width: "25rem",                     // Custom width (increased)
                                my: "center center",               // Position at center
                                at: "center center",               // Align to center
                                of: window,                        // Relative to window
                                offset: "0 0",                     // No offset
                                collision: "fit fit",              // Keep within viewport
                                autoClose: true,                   // Auto close after duration
                                animationDuration: 500,            // Animation duration
                                closeOnBrowserNavigation: true
                            });

                            // Add custom styling with padding to the toast
                            setTimeout(() => {
                                const oToast = document.querySelector(".sapMMessageToast");
                                if (oToast) {
                                    oToast.style.padding = "1.5rem 2rem";
                                    oToast.style.fontSize = "1rem";
                                }
                            }, 50);

                            // Add delay before refresh
                            setTimeout(() => {
                                this.refreshSmartTable();
                            }, 500);
                        },
                        error: (xhr, status, err) => {
                            oBusy.close();
                            console.error("Delete announcement failed:", status, err);
                            console.error("Response:", xhr.responseText);
                            let sErrorMessage = "Failed to delete announcement. Please try again.";
                            if (xhr.responseJSON?.error?.message) {
                                sErrorMessage = xhr.responseJSON.error.message;
                            }
                            sap.m.MessageBox.error(sErrorMessage);
                        }
                    });
                })
                .catch((err) => {
                    oBusy.close();
                    console.error("CSRF token fetch failed:", err);
                    sap.m.MessageBox.error("Failed to initialize request. Please try again.");
                });
        },

        /* ========================================
         * UTILITY METHODS
         * ======================================== */

        _setButtonVisibility: function (oVisibility) {
            const oWizardModel = this.getView().getModel("wizardModel");
            Object.keys(oVisibility).forEach(sKey => {
                oWizardModel.setProperty(`/${sKey}`, oVisibility[sKey]);
            });
        },

        getBaseURL: function () {
            const appId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
            const appPath = appId.replaceAll(".", "/");
            const appModulePath = jQuery.sap.getModulePath(appPath);
            return appModulePath;
        },

        getCurrentUserEmail: async function () {
            try {
                const url = this.getBaseURL() + "/user-api/currentUser";
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

        formatUSDateTime: function (oDate) {
            return formatter.formatUSDateTime(oDate);
        },

        formatDateOnly: function (oDate) {
            return formatter.formatDateOnly(oDate);
        },
        formatDateToDDMMYYYY: function (oDate) {
            return formatter.formatDateToDDMMYYYY(oDate);
        },

        // formatCategoryNames: function (aToTypes) {
        //     const oCategoryModel = this.getView().getModel("categoryModel");
        //     return formatter.formatCategoryNames(aToTypes, oCategoryModel);
        // },

        getModelData: function () {
            return this.getView().getModel().getData();
        },

        refreshTable: function () {
            // this._fetchAnnouncements();
            this.refreshSmartTable();
            const oTable = this.byId("announcementsTable");
            if (oTable) {
                oTable.getBinding("items").refresh();
            }
        }
    });
});