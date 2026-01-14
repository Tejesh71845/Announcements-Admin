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
            this._initializeMainModel();
            this._initWizardModel();
            // 1. Check if typeModel is loaded
            const oTypeModel = this.getOwnerComponent().getModel("typeModel");
            console.log("TypeModel loaded:", !!oTypeModel);

            if (oTypeModel) {
                // 2. Check metadata loading
                oTypeModel.attachMetadataLoaded(() => {
                    console.log("✅ TypeModel metadata loaded successfully");
                });

                oTypeModel.attachMetadataFailed((oEvent) => {
                    console.error("❌ TypeModel metadata failed:", oEvent.getParameters());
                });

                // 3. Log service URL
                console.log("TypeModel Service URL:", oTypeModel.sServiceUrl);
            }

            // 4. Initialize category model
            this._initCategoryModel();
            this._initAnnouncementTypeModel();
            this._editContext = null;
            this._oWizardDialog = null;
            this._oRichTextEditor = null;
        },

        // Add this test function to your controller
        testTypeService: function () {
            const oTypeModel = this.getOwnerComponent().getModel("typeModel");

            // Test 1: Simple read
            oTypeModel.read("/Types", {
                success: (data) => {
                    console.log("✅ Types fetched:", data.results.length);
                    console.log("First type:", data.results[0]);
                },
                error: (err) => {
                    console.error("❌ Error:", err);
                    console.error("Status:", err.statusCode);
                    console.error("Response:", err.responseText);
                }
            });

            // Test 2: Check metadata
            const oMetadata = oTypeModel.getServiceMetadata();
            console.log("Metadata:", oMetadata);
        },

        /**
         * SmartTable initialization - add action column
         */
        onSmartTableInit: function (oEvent) {
            const oSmartTable = oEvent.getSource();
            const oTable = oSmartTable.getTable();

            // Add actions column
            oTable.addColumn(new sap.m.Column({
                hAlign: "Center",
                width: "8rem",
                header: new sap.m.Text({ text: "Actions" })
            }));
        },

        /**
         * Before rebind - add action buttons to template
         */
        onBeforeRebindTable: function (oEvent) {
            const oBindingParams = oEvent.getParameter("bindingParams");
            const oSmartTable = oEvent.getSource();
            const oTable = oSmartTable.getTable();

            if (oBindingParams.parameters) {
                const oTemplate = oBindingParams.parameters.template;

                if (oTemplate) {
                    // Add action buttons cell
                    oTemplate.addCell(new sap.m.HBox({
                        justifyContent: "Center",
                        items: [
                            new sap.m.Button({
                                icon: "sap-icon://edit",
                                type: "Transparent",
                                tooltip: "Edit",
                                press: this.onEditPress.bind(this)
                            }),
                            new sap.m.Button({
                                icon: "sap-icon://delete",
                                type: "Transparent",
                                tooltip: "Delete",
                                press: this.onDeletePress.bind(this)
                            })
                        ]
                    }));
                }
            }
        },

        /* ========================================
         * MODEL INITIALIZATION
         * ======================================== */

        _initializeMainModel: function () {
            const oModel = new JSONModel({ announcements: [] });
            this.getView().setModel(oModel);
        },

        _initAnnouncementTypeModel: function () {
            const oAnnouncementTypeModel = new JSONModel();
            const sModelPath = sap.ui.require.toUrl("com/incture/announcements/model/AnnouncementTypes.json");

            oAnnouncementTypeModel.loadData(sModelPath);
            this.getView().setModel(oAnnouncementTypeModel, "announcementTypeModel");
        },

        _initCategoryModel: function () {
            const oCategoryModel = new sap.ui.model.json.JSONModel({
                category: [],
                idToNameMap: {},
                nameToIdMap: {}
            });
            this.getView().setModel(oCategoryModel, "categoryModel");

            // Use destination with the correct path
            const sUrl = "/JnJ_Workzone_Portal_Destination_Node/odata/v2/type/Types";

            // Show busy indicator
            this.getView().setBusy(true);

            $.ajax({
                url: sUrl,
                method: "GET",
                dataType: "json",
                success: (oData) => {
                    this.getView().setBusy(false);

                    // OData V2 structure: data is in d.results
                    const aEntries = oData.d?.results || [];
                    const aDropdownData = [];
                    const idToName = {};
                    const nameToId = {};

                    aEntries.forEach((entry) => {
                        // Use typeId and name from the response
                        const typeId = entry.typeId;
                        const typeName = entry.name;

                        // Validate data before processing
                        if (typeId && typeName) {
                            aDropdownData.push({
                                key: typeId,
                                text: typeName
                            });
                            idToName[typeId] = typeName;
                            nameToId[typeName] = typeId;
                        } else {
                            console.warn("Invalid type entry:", entry);
                        }
                    });

                    oCategoryModel.setProperty("/category", aDropdownData);
                    oCategoryModel.setProperty("/idToNameMap", idToName);
                    oCategoryModel.setProperty("/nameToIdMap", nameToId);

                    console.log("Categories loaded:", aDropdownData.length);

                    // Fetch announcements after categories are loaded
                    this._fetchAnnouncements();
                },
                error: (jqXHR, textStatus, errorThrown) => {
                    this.getView().setBusy(false);

                    // Enhanced error logging
                    console.error("Failed to fetch category data:", {
                        status: jqXHR.status,
                        statusText: textStatus,
                        error: errorThrown,
                        response: jqXHR.responseText
                    });

                    // Parse error response
                    let errorMessage = "Failed to fetch category data";
                    try {
                        if (jqXHR.responseText) {
                            const errorObj = JSON.parse(jqXHR.responseText);
                            errorMessage = errorObj.error?.message?.value || errorObj.error?.message || errorMessage;
                        }
                    } catch (e) {
                        console.error("Error parsing error response:", e);
                        errorMessage = `${errorMessage} (Status: ${jqXHR.status})`;
                    }

                    // Show detailed error to user
                    sap.m.MessageBox.error(errorMessage, {
                        title: "Error Loading Categories",
                        details: `${textStatus || "Internal Server Error"}\nStatus: ${jqXHR.status}`
                    });
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
                showGroupInsert: true,
                value: sDescription,
                valueState: "{wizardModel>/descriptionValueState}",           // Add this
                valueStateText: "{wizardModel>/descriptionValueStateText}",   // Add this
                ready: function () {
                    this.addButtonGroup("styles").addButtonGroup("table");
                },
                change: (oEvent) => {
                    const sValue = oEvent.getParameter("newValue");
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

        _fetchAnnouncements: function () {
            const sUrl = "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/Announcements?$expand=toTypes($expand=type)";

            $.ajax({
                url: sUrl,
                method: "GET",
                dataType: "json",
                success: (oData) => {
                    const oCategoryModel = this.getView().getModel("categoryModel");
                    const idToNameMap = oCategoryModel.getProperty("/idToNameMap") || {};
                    const aEntries = oData.value || [];

                    // Filter out inactive announcements
                    const aActiveEntries = aEntries.filter(entry => entry.isActive !== false);

                    const aMapped = aActiveEntries.map((entry) => {
                        const aTypeIds = (entry.toTypes || [])
                            .map(item => item.type?.typeId)
                            .filter(Boolean);

                        const aCategoryNames = aTypeIds
                            .map(id => idToNameMap[id])
                            .filter(Boolean);

                        return {
                            id: entry.announcementId,
                            title: entry.title,
                            category: aCategoryNames.join(", ") || "N/A",
                            description: entry.description,
                            announcementType: entry.announcementType,
                            createdOn: entry.createdAt,
                            createdBy: entry.createdBy,
                            modifiedOn: entry.modifiedAt,
                            modifiedBy: entry.modifiedBy,
                            publishedAt: entry.publishedAt,           // NEW
                            publishedBy: entry.publishedBy,           // NEW
                            announcementStatus: entry.announcementStatus || "DRAFT",  // NEW
                            publishStatus: entry.publishStatus,       // NEW
                            startAnnouncement: entry.startAnnouncement, // NEW
                            endAnnouncement: entry.endAnnouncement,   // NEW
                            typeId: aTypeIds,
                            isActive: entry.isActive
                        };
                    });

                    // Sort announcements by createdOn date in descending order
                    aMapped.sort((a, b) => {
                        const dateA = a.createdOn ? new Date(a.createdOn).getTime() : 0;
                        const dateB = b.createdOn ? new Date(b.createdOn).getTime() : 0;
                        return dateB - dateA;
                    });

                    this.getView().getModel().setProperty("/announcements", aMapped);
                },
                error: (xhr, status, err) => {
                    console.error("Failed to fetch announcements:", status, err);
                    sap.m.MessageBox.error("Unable to load announcements.");
                }
            });
        },
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
                publishOption: oEditData.publishStatus === "PUBLISH_NOW" ? "PUBLISH" :
                    oEditData.publishStatus === "PUBLISH_LATER" ? "UNPUBLISH" : "",
                publishStartDate: oEditData.startAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.startAnnouncement)) : "",
                publishEndDate: oEditData.endAnnouncement ?
                    this._formatDateToValue(new Date(oEditData.endAnnouncement)) : "",
                showDatePickers: bHasPublishData,
                startDateEnabled: oEditData.publishStatus === "PUBLISH_LATER",
                endDateEnabled: bHasPublishData,

                publishEnabled: bHasPublishData,
                cancelEnabled: true,

                // Store original publish data for comparison
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
            oSevenDaysLater.setDate(oSevenDaysLater.getDate() + 7);

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

            // FIXED: Show Cancel, Draft, Submit buttons when entering create step
            if (bIsEditMode) {
                this._setButtonVisibility({
                    showCancelButton: true,
                    showDraftButton: true,
                    showSubmitButton: true,
                    showReviewButton: false,
                    showPublishButton: false,
                    showUpdateButton: false,
                    showResetButton: false
                });
            } else {
                this._setButtonVisibility({
                    showCancelButton: true,
                    showDraftButton: true,
                    showSubmitButton: true,
                    showReviewButton: false,
                    showPublishButton: false,
                    showUpdateButton: false,
                    showResetButton: false
                });
            }
        },

        onSubmitPress: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");

            if (sCurrentFlow === "BULK") {
                // Validate bulk table data before submitting
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

                // If validation passes, confirm and submit
                MessageBox.confirm(
                    `Are you sure you want to submit ${bulkData.length} announcement(s)?`,
                    {
                        title: "Confirm Submit",
                        actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                        emphasizedAction: MessageBox.Action.YES,
                        onClose: (oAction) => {
                            if (oAction === MessageBox.Action.YES) {
                                this._handleBulkSubmit();
                            }
                        }
                    }
                );
            } else {
                // Single flow validation
                if (!this._validateMandatoryFields()) {
                    MessageToast.show("Please complete all required fields");
                    return;
                }

                if (!this._validatePublishOptions()) {
                    return;
                }

                const bIsEditMode = oWizardModel.getProperty("/isEditMode");
                const sConfirmMessage = bIsEditMode
                    ? "Are you sure you want to submit these changes?"
                    : "Are you sure you want to submit this announcement?";

                MessageBox.confirm(sConfirmMessage, {
                    title: "Confirm Submit",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.YES,
                    onClose: (oAction) => {
                        if (oAction === MessageBox.Action.YES) {
                            this._handleSubmitAnnouncement();
                        }
                    }
                });
            }
        },

        _validateMandatoryFields: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();

            return sTitle && aAnnouncementType.length > 0 && aCategory.length > 0 && sDescription;
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

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    let announcementStatus, startAnnouncement, endAnnouncement;
                    const currentDateTime = new Date().toISOString();

                    if (sPublishOption === "PUBLISH") {
                        // Publish Now: Status = PUBLISHED
                        announcementStatus = "PUBLISHED";
                        startAnnouncement = currentDateTime;
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    } else {
                        // Publish Later: Status = TO_BE_PUBLISHED
                        announcementStatus = "TO_BE_PUBLISHED";
                        startAnnouncement = new Date(sStartDate).toISOString();
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    }

                    const oPayload = {
                        data: [{
                            title: sTitle,
                            description: sDescription,
                            announcementType: sAnnouncementType,
                            isRead: false,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            announcementStatus: announcementStatus,
                            publishedBy: sUserEmail,
                            publishedAt: startAnnouncement, // Use start date as published date
                            toTypes: aTypeIds.map(typeId => ({
                                type: { typeId: typeId }
                            }))
                        }]
                    };

                    const oBusy = new sap.m.BusyDialog({ text: "Submitting announcement..." });
                    oBusy.open();

                    $.ajax({
                        url: "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/bulkCreateAnnouncements",
                        method: "POST",
                        contentType: "application/json",
                        dataType: "json",
                        data: JSON.stringify(oPayload),
                        success: (oResponse) => {
                            oBusy.close();
                            this._oWizardDialog.close();
                            sap.m.MessageToast.show(`Announcement '${sTitle}' submitted successfully!`);
                            this._fetchAnnouncements();
                        },
                        error: (xhr, status, err) => {
                            oBusy.close();
                            console.error("Submit failed:", status, err);
                            let sErrorMessage = "Failed to submit announcement.";
                            if (xhr.responseJSON?.error?.message) {
                                sErrorMessage = xhr.responseJSON.error.message;
                            }
                            sap.m.MessageBox.error(sErrorMessage);
                        }
                    });
                });
        },
        /* ========================================
         * INPUT VALIDATION & CHANGE HANDLERS
         * ======================================== */

        _validateRichTextDescription: function (oWizardModel) {
            // Get raw HTML from wizard model
            const sDescriptionHTML = oWizardModel.getProperty("/description") || "";

            // Remove HTML tags and check plain text length
            const sPlainText = sDescriptionHTML.replace(/<[^>]*>/g, "").trim();
            const bValid = sPlainText.length > 0;

            // Apply valueState and valueStateText
            this._setValidationState(oWizardModel, "description", bValid, "Description is required");

            // Highlight/Remove red border manually for RichTextEditor
            const oContainer = this.byId("richTextContainer");
            if (oContainer) {
                if (!bValid) {
                    oContainer.addStyleClass("richTextError");
                } else {
                    oContainer.removeStyleClass("richTextError");
                }
            }

            // Fetch other fields for overall step validity
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];

            const bOverallValid =
                sTitle.length > 0 &&
                aAnnouncementType.length > 0 &&
                aCategory.length > 0 &&
                bValid;

            // Update wizard step validation
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
                sValue = sValue.replace(/[^a-zA-Z0-9 ]/g, "");

                const iMaxLength = oSource.getMaxLength ? oSource.getMaxLength() : 500;
                if (sValue.length > iMaxLength) {
                    sValue = sValue.substring(0, iMaxLength);
                }

                if (sValue !== oSource.getValue()) {
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

            const bOverallValid = sTitle && aAnnouncementType.length > 0 && aCategory.length > 0 && sDescription;

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
            const bOverallValid = sTitle && bValid && aCategory.length > 0 && sDescription;
            oWizardModel.setProperty("/singleCreateStepValidated", !!bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(!!bOverallValid);
        },

        _validateMultiCategory: function (oWizardModel) {
            const aCategories = oWizardModel.getProperty("/category") || [];
            const bValid = Array.isArray(aCategories) && aCategories.length > 0;
            this._setValidationState(oWizardModel, "category", bValid, "At least one category is required");

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();
            const bOverallValid = sTitle && aAnnouncementType.length > 0 && bValid && sDescription;
            oWizardModel.setProperty("/singleCreateStepValidated", !!bOverallValid);

            const oStep = this.byId("singleCreateStep");
            if (oStep) oStep.setValidated(!!bOverallValid);
        },

        _handleResetButtonVisibility: function (oWizardModel) {
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");
            const sTitle = oWizardModel.getProperty("/title") || "";
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = oWizardModel.getProperty("/description") || "";

            if (bIsEditMode) {
                const sOriginalTitle = oWizardModel.getProperty("/originalTitle") || "";
                const aOriginalAnnouncementType = oWizardModel.getProperty("/originalAnnouncementType") || [];
                const aOriginalCategory = oWizardModel.getProperty("/originalCategory") || [];
                const sOriginalDescription = oWizardModel.getProperty("/originalDescription") || "";

                const bHasChanged = (sTitle !== sOriginalTitle) ||
                    (JSON.stringify(aAnnouncementType.sort()) !== JSON.stringify(aOriginalAnnouncementType.sort())) ||
                    (JSON.stringify(aCategory.sort()) !== JSON.stringify(aOriginalCategory.sort())) ||
                    (sDescription !== sOriginalDescription);

                oWizardModel.setProperty("/showResetButton", bHasChanged);
            } else {
                const bUserHasStartedTyping = sTitle.length > 0 ||
                    aAnnouncementType.length > 0 ||
                    aCategory.length > 0 ||
                    sDescription.length > 0;

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
                // Publish Now: Start Date disabled, End Date enabled
                const sTodayValue = this._formatDateToValue(oToday);

                const oEndDate = new Date(oToday);
                oEndDate.setDate(oEndDate.getDate() + 7);
                const sEndDateValue = this._formatDateToValue(oEndDate);

                oWizardModel.setProperty("/startDateEnabled", false);
                oWizardModel.setProperty("/endDateEnabled", true);
                oWizardModel.setProperty("/publishStartDate", sTodayValue);
                oWizardModel.setProperty("/publishEndDate", sEndDateValue);
                oWizardModel.setProperty("/publishStartDateValueState", "None");
                oWizardModel.setProperty("/publishStartDateValueStateText", "");
                oWizardModel.setProperty("/publishEndDateValueState", "None");
                oWizardModel.setProperty("/publishEndDateValueStateText", "");


                oWizardModel.setProperty("/minEndDate", oToday);

                oWizardModel.setProperty("/draftEnabled", true);


                // Force DatePicker to update display format
                const oStartDatePicker = this.byId("publishStartDatePicker");
                if (oStartDatePicker) {
                    oStartDatePicker.setValue(sTodayValue);
                }
            } else {
                // Publish Later: Both enabled, start date is tomorrow, min date is tomorrow
                const oTomorrow = new Date(oToday);
                oTomorrow.setDate(oTomorrow.getDate() + 1);
                const sTomorrowValue = this._formatDateToValue(oTomorrow);

                // End date is tomorrow + 7 days
                const oEndDate = new Date(oTomorrow);
                oEndDate.setDate(oEndDate.getDate() + 7);
                const sEndDateValue = this._formatDateToValue(oEndDate);

                oWizardModel.setProperty("/startDateEnabled", true);
                oWizardModel.setProperty("/endDateEnabled", true);
                oWizardModel.setProperty("/publishStartDate", sTomorrowValue);
                oWizardModel.setProperty("/publishEndDate", sEndDateValue);

                // Set minimum date to tomorrow for Publish Later
                oWizardModel.setProperty("/minPublishDate", oTomorrow);
                oWizardModel.setProperty("/minEndDate", oTomorrow);

                oWizardModel.setProperty("/draftEnabled", false);
            }

            // Enable Cancel and Publish buttons
            oWizardModel.setProperty("/cancelEnabled", true);
            oWizardModel.setProperty("/publishEnabled", true);
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

            // Auto-populate End Date (Start Date + 7 days)
            const oEndDate = new Date(oSelected);
            oEndDate.setDate(oEndDate.getDate() + 7);
            const sEndDateValue = this._formatDateToValue(oEndDate);

            oWizardModel.setProperty("/publishEndDate", sEndDateValue);
            oWizardModel.setProperty("/publishEndDateValueState", "None");
            oWizardModel.setProperty("/publishEndDateValueStateText", "");

            // Update minimum date for End Date picker to selected start date
            oWizardModel.setProperty("/minEndDate", oSelected);
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

            // Validate End Date is after or equal to Start Date
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            if (sStartDate) {
                const oStartDate = new Date(sStartDate);
                oStartDate.setHours(0, 0, 0, 0);
                if (oSelected < oStartDate) { // **CHANGED: from <= to <**
                    oWizardModel.setProperty("/publishEndDateValueState", "Error");
                    oWizardModel.setProperty("/publishEndDateValueStateText", "End Date must be on or after Start Date");
                    return;
                }
            }

            oWizardModel.setProperty("/publishEndDateValueState", "None");
            oWizardModel.setProperty("/publishEndDateValueStateText", "");
        },

        _validateSingleCreateStep: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementType = oWizardModel.getProperty("/announcementType") || [];
            const aCategory = oWizardModel.getProperty("/category") || [];
            const sDescription = (oWizardModel.getProperty("/description") || "").trim();

            const bTitleValid = sTitle.length > 0;
            const bAnnouncementTypeValid = aAnnouncementType.length > 0;
            const bCategoryValid = aCategory.length > 0;
            const bDescriptionValid = sDescription.length > 0;

            const bOverallValid = bTitleValid && bAnnouncementTypeValid && bCategoryValid && bDescriptionValid;

            this._setValidationState(oWizardModel, "title", bTitleValid, "Title is required");
            this._setValidationState(oWizardModel, "announcementType", bAnnouncementTypeValid, "Announcement Type is required");
            this._setValidationState(oWizardModel, "category", bCategoryValid, "Category is required");
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
            const oBindingContext = oEvent.getSource().getBindingContext();
            const oData = oBindingContext.getObject();

            this._editContext = {
                data: oData,
                path: oBindingContext.getPath(),
                isEdit: true
            };

            this._openWizardForEdit(oData);
        },

        _openWizardForEdit: function (oEditData) {
            this._loadWizardDialog()
                .then(() => {
                    // Initialize wizard model with edit data
                    this._initWizardModelForEdit(oEditData);

                    // Reset wizards to clean state
                    this._resetWizards();

                    // Open dialog first - this is critical for DOM rendering
                    this._oWizardDialog.open();

                    // Wait for dialog to fully render before setting up wizard
                    setTimeout(() => {
                        this._setupSingleWizardForEdit();
                    }, 150);
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
                return;
            }

            // Hide selection step (Edit mode)
            oSelectionStep.setVisible(false);
            oSelectionStep.setValidated(true);

            // Validate create step
            oCreateStep.setValidated(true);

            // Set current step
            oWizard.setCurrentStep(oCreateStep);

            // FIXED: Set footer buttons for edit mode - show Draft and Submit
            this._setButtonVisibility({
                showCancelButton: true,
                showDraftButton: true,
                showSubmitButton: true,
                showReviewButton: false,
                showPublishButton: false,
                showUpdateButton: false,
                showResetButton: false
            });

            // Force layout + initialize RichTextEditor
            setTimeout(() => {
                oWizard.invalidate();

                setTimeout(() => {
                    this._initRichTextEditor("richTextContainer");
                }, 100);
            }, 150);
        },


        /* ========================================
         * SUBMIT HANDLERS
         * ======================================== */

        onPublishPress: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sCurrentFlow = oWizardModel.getProperty("/currentFlow");
            const bIsEditMode = oWizardModel.getProperty("/isEditMode");

            // Validate publish options for BOTH flows
            const bValidPublish = this._validatePublishOptions();
            if (!bValidPublish) {
                return;
            }

            // Show confirmation dialog before publishing
            const sConfirmMessage = bIsEditMode
                ? "Are you sure you want to publish these changes?"
                : (sCurrentFlow === "BULK"
                    ? "Are you sure you want to publish these announcements?"
                    : "Are you sure you want to publish this announcement?");

            MessageBox.confirm(sConfirmMessage, {
                title: "Confirm Publish",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.YES,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        // Call submit handlers
                        if (sCurrentFlow === "SINGLE") {
                            bIsEditMode ? this._handleEditSubmit() : this._handleSingleSubmit();
                        } else if (sCurrentFlow === "BULK") {
                            this._handleBulkSubmit();
                        }
                    }
                }
            });
        },

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

        _handleDraftSubmit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aTypeIds = oWizardModel.getProperty("/category") || [];

            // No validation - allow saving with empty fields
            const sDescription = this._stripHtmlTags(sDescriptionHTML);

            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];
            const aSelectedTypes = aAnnouncementTypeKeys
                .map(key => aTypeList.find(t => t.key === key)?.text)
                .filter(Boolean);
            const sAnnouncementType = aSelectedTypes.join(",");

            // Only get publish dates if Publish Now is selected
            const sPublishOption = oWizardModel.getProperty("/publishOption");
            let startAnnouncement = null;
            let endAnnouncement = null;

            if (sPublishOption === "PUBLISH") {
                const sEndDate = oWizardModel.getProperty("/publishEndDate");
                if (sEndDate) {
                    startAnnouncement = new Date().toISOString();
                    const oEndDate = new Date(sEndDate);
                    oEndDate.setDate(oEndDate.getDate() + 1);
                    endAnnouncement = oEndDate.toISOString();
                }
            }

            const oPayload = {
                data: [{
                    title: sTitle,
                    description: sDescription,
                    announcementType: sAnnouncementType,
                    isRead: false,
                    announcementStatus: "DRAFT",
                    startAnnouncement: startAnnouncement,
                    endAnnouncement: endAnnouncement,
                    publishedBy: null,
                    publishedAt: null,
                    toTypes: aTypeIds.map(typeId => ({
                        type: { typeId: typeId }
                    }))
                }]
            };

            const oBusy = new sap.m.BusyDialog({ text: "Saving as draft..." });
            oBusy.open();

            $.ajax({
                url: "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/bulkCreateAnnouncements",
                method: "POST",
                contentType: "application/json",
                dataType: "json",
                data: JSON.stringify(oPayload),
                success: (oResponse) => {
                    oBusy.close();
                    this._oWizardDialog.close();
                    sap.m.MessageToast.show("Announcement saved as draft successfully!");
                    this._fetchAnnouncements();
                },
                error: (xhr, status, err) => {
                    oBusy.close();
                    console.error("Save as draft failed:", status, err);
                    sap.m.MessageBox.error("Failed to save as draft.");
                }
            });
        },
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

            if (sOption === "UNPUBLISH") {
                // Validate End Date is on or after Start Date
                const oStartDate = new Date(sStartDate);
                oStartDate.setHours(0, 0, 0, 0);
                if (oEndDate < oStartDate) {
                    oWizardModel.setProperty("/publishEndDateValueState", "Error");
                    oWizardModel.setProperty("/publishEndDateValueStateText", "End Date must be on or after Start Date");
                    MessageBox.error("End Date must be on or after Start Date.");
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

        _handleSingleSubmit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const oCategoryModel = this.getView().getModel("categoryModel");

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aTypeIds = oWizardModel.getProperty("/category") || [];

            // Get publish option and dates
            const sPublishOption = oWizardModel.getProperty("/publishOption");
            const sStartDate = oWizardModel.getProperty("/publishStartDate");
            const sEndDate = oWizardModel.getProperty("/publishEndDate");

            if (!sTitle || !sDescriptionHTML || aAnnouncementTypeKeys.length === 0 || aTypeIds.length === 0) {
                sap.m.MessageToast.show("Please fill in all required fields");
                return; F
            }

            const sDescription = this._stripHtmlTags(sDescriptionHTML);

            const oAnnouncementTypeModel = this.getView().getModel("announcementTypeModel");
            const aTypeList = oAnnouncementTypeModel.getProperty("/types") || [];

            const aSelectedTypes = aAnnouncementTypeKeys
                .map(key => aTypeList.find(t => t.key === key)?.text)
                .filter(Boolean);

            const sAnnouncementType = aSelectedTypes.join(",");

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    // Determine announcement status and dates
                    let announcementStatus, startAnnouncement, endAnnouncement, publishStatus;
                    const currentDateTime = new Date().toISOString();

                    if (sPublishOption === "PUBLISH") {
                        // Publish Now
                        announcementStatus = "PUBLISHED";
                        publishStatus = "PUBLISH_NOW";
                        startAnnouncement = currentDateTime;
                        // Add 1 day to end date to include the full day
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    } else {
                        // Publish Later
                        announcementStatus = "TO_BE_PUBLISHED";
                        publishStatus = "PUBLISH_LATER";
                        startAnnouncement = new Date(sStartDate).toISOString();
                        // Add 1 day to end date to include the full day
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    }

                    const oPayload = {
                        data: [
                            {
                                title: sTitle,
                                description: sDescription,
                                announcementType: sAnnouncementType,
                                isRead: false,
                                startAnnouncement: startAnnouncement,
                                endAnnouncement: endAnnouncement,
                                publishStatus: publishStatus,
                                announcementStatus: announcementStatus,
                                publishedBy: sUserEmail,
                                publishedAt: currentDateTime,
                                toTypes: aTypeIds.map(typeId => ({
                                    type: { typeId: typeId }
                                }))
                            }
                        ]
                    };

                    const oBusy = new sap.m.BusyDialog({ text: "Creating announcement..." });
                    oBusy.open();

                    $.ajax({
                        url: "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/bulkCreateAnnouncements",
                        method: "POST",
                        contentType: "application/json",
                        dataType: "json",
                        data: JSON.stringify(oPayload),

                        success: (oResponse) => {
                            oBusy.close();
                            this._oWizardDialog.close();

                            const sMessage = announcementStatus === "PUBLISHED"
                                ? `Announcement '${sTitle}' published successfully!`
                                : `Announcement '${sTitle}' scheduled for publication!`;

                            sap.m.MessageToast.show(sMessage);
                            this._fetchAnnouncements();
                        },
                        error: (xhr, status, err) => {
                            oBusy.close();
                            console.error("Create announcement failed:", status, err);
                            let sErrorMessage = "Failed to create announcement. Please try again.";
                            if (xhr.responseJSON?.error?.message) {
                                sErrorMessage = xhr.responseJSON.error.message;
                            }
                            sap.m.MessageBox.error(sErrorMessage);
                        }
                    });
                })
                .catch((err) => {
                    sap.m.MessageBox.error("Failed to get current user: " + err.message);
                });
        },

        _handleEditSubmit: function () {
            const oWizardModel = this.getView().getModel("wizardModel");
            const oCategoryModel = this.getView().getModel("categoryModel");

            const sTitle = (oWizardModel.getProperty("/title") || "").trim();
            const aAnnouncementTypeKeys = oWizardModel.getProperty("/announcementType") || [];
            const aCategories = oWizardModel.getProperty("/category") || [];
            const sDescriptionHTML = (oWizardModel.getProperty("/description") || "").trim();
            const sEditId = oWizardModel.getProperty("/editId");

            // Get publish option and dates
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

            const oBusyDialog = new sap.m.BusyDialog({ text: "Updating announcement..." });
            oBusyDialog.open();

            this.getCurrentUserEmail()
                .then((sUserEmail) => {
                    // Determine announcement status and dates
                    let announcementStatus, startAnnouncement, endAnnouncement, publishStatus;
                    const currentDateTime = new Date().toISOString();

                    if (sPublishOption === "PUBLISH") {
                        // Publish Now
                        announcementStatus = "PUBLISHED";
                        publishStatus = "PUBLISH_NOW";
                        startAnnouncement = currentDateTime;
                        // Add 1 day to end date to include the full day
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    } else {
                        // Publish Later
                        announcementStatus = "TO_BE_PUBLISHED";
                        publishStatus = "PUBLISH_LATER";
                        startAnnouncement = new Date(sStartDate).toISOString();
                        // Add 1 day to end date to include the full day
                        const oEndDate = new Date(sEndDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();
                    }

                    const oPayload = {
                        data: {
                            announcementId: sEditId,
                            title: sTitle,
                            description: sDescription,
                            announcementType: sAnnouncementType,
                            isRead: false,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishStatus: publishStatus,
                            announcementStatus: announcementStatus,
                            publishedBy: sUserEmail,
                            publishedAt: currentDateTime,
                            toTypes: aCategories.map(typeId => ({
                                type: { typeId: typeId }
                            }))
                        }
                    };

                    const sUrl = `/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/updateAnnouncement`;

                    $.ajax({
                        url: sUrl,
                        method: "POST",
                        contentType: "application/json",
                        dataType: "json",
                        data: JSON.stringify(oPayload),

                        success: (oResponse) => {
                            oBusyDialog.close();

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
                            this._fetchAnnouncements();
                        },
                        error: (xhr, status, err) => {
                            oBusyDialog.close();
                            console.error("Failed to update announcement:", status, err);
                            let sErrorMessage = "Failed to update announcement. Please try again.";
                            if (xhr.responseJSON?.error?.message) {
                                sErrorMessage = xhr.responseJSON.error.message;
                            }
                            sap.m.MessageBox.error(sErrorMessage);
                        }
                    });
                })
                .catch((err) => {
                    oBusyDialog.close();
                    sap.m.MessageBox.error("Failed to get current user: " + err.message);
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

            // Validate all rows in the table
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

                        // Determine publish status based on start date
                        const oStartDate = new Date(item.startDate);
                        oStartDate.setHours(0, 0, 0, 0);

                        let publishStatus, announcementStatus, startAnnouncement, endAnnouncement;
                        const currentDateTime = new Date().toISOString();

                        if (oStartDate.getTime() === oToday.getTime()) {
                            // Start date is today - PUBLISH NOW
                            publishStatus = "PUBLISH_NOW";
                            announcementStatus = "PUBLISHED";
                            startAnnouncement = currentDateTime;
                        } else {
                            // Start date is in future - PUBLISH LATER
                            publishStatus = "PUBLISH_LATER";
                            announcementStatus = "TO_BE_PUBLISHED";
                            startAnnouncement = new Date(item.startDate).toISOString();
                        }

                        // Add 1 day to end date to include the full day
                        const oEndDate = new Date(item.endDate);
                        oEndDate.setDate(oEndDate.getDate() + 1);
                        endAnnouncement = oEndDate.toISOString();

                        aPayloads.push({
                            title: item.title.trim(),
                            description: item.description.trim(),
                            announcementType: sAnnouncementType,
                            isRead: false,
                            announcementStatus: announcementStatus,
                            publishStatus: publishStatus,
                            startAnnouncement: startAnnouncement,
                            endAnnouncement: endAnnouncement,
                            publishedBy: sUserEmail,
                            publishedAt: currentDateTime,
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

                    const oBusyDialog = new sap.m.BusyDialog({
                        text: `Publishing ${aPayloads.length} announcement(s)...`
                    });
                    oBusyDialog.open();

                    const oPayload = {
                        data: aPayloads
                    };

                    $.ajax({
                        url: "/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/bulkCreateAnnouncements",
                        method: "POST",
                        contentType: "application/json",
                        dataType: "json",
                        data: JSON.stringify(oPayload),
                        success: (oResponse) => {
                            oBusyDialog.close();

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
                            this._fetchAnnouncements();
                        },
                        error: (xhr, status, err) => {
                            oBusyDialog.close();
                            console.error("Bulk upload failed:", status, err);

                            let sErrorMessage = "Bulk upload failed. Please check the data and try again.";
                            if (xhr.responseJSON?.error?.message) {
                                sErrorMessage = xhr.responseJSON.error.message;
                            }

                            sap.m.MessageBox.error(sErrorMessage);
                        }
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
                const sOriginalTitle = oWizardModel.getProperty("/originalTitle") || "";
                const aOriginalAnnouncementType = oWizardModel.getProperty("/originalAnnouncementType") || [];
                const aOriginalCategory = oWizardModel.getProperty("/originalCategory") || [];
                const sOriginalDescription = oWizardModel.getProperty("/originalDescription") || "";

                oWizardModel.setProperty("/title", sOriginalTitle);
                oWizardModel.setProperty("/announcementType", aOriginalAnnouncementType);
                oWizardModel.setProperty("/category", aOriginalCategory);
                oWizardModel.setProperty("/description", sOriginalDescription);
                // UPDATE RICHTEXTEDITOR VALUE
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue(sOriginalDescription);
                }
            } else {
                oWizardModel.setProperty("/title", "");
                oWizardModel.setProperty("/announcementType", []);
                oWizardModel.setProperty("/category", []);
                oWizardModel.setProperty("/description", "");
                // UPDATE RICHTEXTEDITOR VALUE
                if (this._oRichTextEditor) {
                    this._oRichTextEditor.setValue("");
                }
            }

            this._clearValidationErrors(oWizardModel);
            oWizardModel.setProperty("/singleCreateStepValidated", bIsEditMode);
            oWizardModel.setProperty("/showResetButton", false);

            if (bIsEditMode) {
                this._setButtonVisibility({
                    showReviewButton: true,
                    showUpdateButton: false,
                    showSubmitButton: false,
                    showResetButton: false
                });
            } else {
                this._setButtonVisibility({
                    showReviewButton: true,
                    showUpdateButton: false,
                    showSubmitButton: false,
                    showResetButton: false
                });
            }

            const oStep = this.byId("singleCreateStep");
            if (oStep && bIsEditMode) {
                oStep.setValidated(true);
            }

            MessageToast.show(bIsEditMode ? "Reset to original values" : "Form has been reset");
        },

        /* ========================================
         * DELETE FUNCTIONALITY
         * ======================================== */

        onDeletePress: function (oEvent) {
            const oBindingContext = oEvent.getSource().getBindingContext();
            const oData = oBindingContext.getObject();

            const sAnnouncementId = oData.id;
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

        _deleteItem: function (sAnnouncementId, sTitle) {
            const oBusyDialog = new sap.m.BusyDialog({
                text: "Deleting announcement..."
            });
            oBusyDialog.open();

            const sUrl = `/JnJ_Workzone_Portal_Destination_Java/odata/v4/AnnouncementService/Announcements(announcementId=${sAnnouncementId})`;

            $.ajax({
                url: sUrl,
                method: "DELETE",
                success: () => {
                    oBusyDialog.close();

                    const oModel = this.getView().getModel();
                    const aAnnouncements = oModel.getProperty("/announcements") || [];
                    const iIndex = aAnnouncements.findIndex(item => item.id === sAnnouncementId);

                    if (iIndex !== -1) {
                        aAnnouncements.splice(iIndex, 1);
                        oModel.setProperty("/announcements", aAnnouncements);
                    }

                    MessageToast.show(`Announcement '${sTitle}' deleted successfully!`);
                },
                error: (xhr, status, err) => {
                    oBusyDialog.close();
                    console.error("Failed to delete announcement:", status, err);

                    let sErrorMessage = "Failed to delete announcement. Please try again.";
                    if (xhr.responseJSON && xhr.responseJSON.error && xhr.responseJSON.error.message) {
                        sErrorMessage = xhr.responseJSON.error.message;
                    }

                    MessageBox.error(sErrorMessage);
                }
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

        getModelData: function () {
            return this.getView().getModel().getData();
        },

        refreshTable: function () {
            this._fetchAnnouncements();
            const oTable = this.byId("announcementsTable");
            if (oTable) {
                oTable.getBinding("items").refresh();
            }
        }
    });
});