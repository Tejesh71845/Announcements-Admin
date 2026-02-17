sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "com/incture/announcements/utils/formatter"
], (Controller, MessageToast, MessageBox, formatter) => {
    "use strict";

    return Controller.extend("com.incture.announcements.controller.Announcement", {

        formatter: formatter,

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

            var oAnnouncementsSmartTable = this.getView().byId("idAnnouncementsSmartTbl");
            oAnnouncementsSmartTable.setModel(oAnnouncementModel);

            setTimeout(() => {
                this.refreshSmartTable();
            }, 300);
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
            const oSmartTable = this.byId("idAnnouncementsSmartTbl");
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


        _formatDateToDisplay: function (oDate) {
            return formatter.formatDateToDisplay(oDate);
        },


        formatStatusState: function (sStatus) {
            return formatter.formatStatusState(sStatus);
        },

        formatStatusText: function (sStatus) {
            return formatter.formatStatusText(sStatus);
        },

        onCreateSidebarPress: function () {
            this._router.navTo("CreateSidebarAnnouncement");
        },

        onCreateBannerPress: function () {
            this._router.navTo("CreateBannerAnnouncement");
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

            const oData = oBindingContext.getObject();
            const sAnnouncementId = oData.announcementId;
            const sAnnouncementType = oData.announcementType;

            if (!sAnnouncementId) {
                MessageBox.error("Unable to get announcement ID.");
                return;
            }

            // Navigate based on announcement type
            if (sAnnouncementType === "Banner") {
                // Navigate to Banner page with announcement ID
                this._router.navTo("CreateBannerAnnouncement", {
                    announcementId: sAnnouncementId
                });
            } else if (sAnnouncementType === "Sidebar" || sAnnouncementType === "Sidebar (Popup)") {
                // Navigate to Sidebar page with announcement ID
                this._router.navTo("CreateSidebarAnnouncement", {
                    announcementId: sAnnouncementId
                });
            } else {
                MessageBox.error("Unknown announcement type: " + sAnnouncementType + ". Cannot edit.");
            }
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
                                const oToast = document.querySelector(".messageToastCss");
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

        formatUSDateTime: function (oDate) {
            return formatter.formatUSDateTime(oDate);
        },

        formatDateOnly: function (oDate) {
            return formatter.formatDateOnly(oDate);
        },
        formatDateToDDMMYYYY: function (oDate) {
            return formatter.formatDateToDDMMYYYY(oDate);
        },

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