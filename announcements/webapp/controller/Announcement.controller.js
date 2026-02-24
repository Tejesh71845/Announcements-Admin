sap.ui.define([
    "com/incture/announcements/controller/BaseController",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator"
], (BaseController, MessageToast, MessageBox, JSONModel, Filter, FilterOperator) => {
    "use strict";

    const STATUS = {
        PUBLISHED: "PUBLISHED",
        TO_BE_PUBLISHED: "TO_BE_PUBLISHED",
        EXPIRED: "EXPIRED",
        INVALID: "INVALID"
    };

    const FILTER_KEY = {
        ALL: "all",
        PUBLISHED: "published",
        SCHEDULED: "scheduled",
        EXPIRED: "expired",
        INVALID: "invalid"
    };

    const FILTER_KEY_TO_STATUS = {
        published: STATUS.PUBLISHED,
        scheduled: STATUS.TO_BE_PUBLISHED,
        expired: STATUS.EXPIRED,
        invalid: STATUS.INVALID
    };

    return BaseController.extend("com.incture.announcements.controller.Announcement", {
        onInit: function () {
            const oComponent = this.getOwnerComponent();
            this._router = oComponent.getRouter();
            this._router.getRoute("Announcement").attachPatternMatched(this._handleRouteMatched, this);

            this._initAnnouncementCountModel();
            this._sActiveFilterKey = FILTER_KEY.ALL;
            this._sSearchQuery = "";
        },

        _initAnnouncementCountModel: function () {
            const oCountModel = new JSONModel({
                activeFilter: FILTER_KEY.ALL,
                allCount: 0,
                publishedCount: 0,
                scheduledCount: 0,
                expiredCount: 0,
                invalidCount: 0
            });
            this.getOwnerComponent().setModel(oCountModel, "announcementCountModel");
        },

        _handleRouteMatched: async function () {
            const oAnnouncementModel = this.getOwnerComponent().getModel("announcementModel");
            const oSmartTable = this.getView().byId("idAnnouncementsSmartTbl");
            oSmartTable.setModel(oAnnouncementModel);
            setTimeout(() => { this.refreshSmartTable(); }, 300);
        },

        onBeforeRebindTable: function (oEvent) {
            const oBindingParams = oEvent.getParameter("bindingParams");

            // Always filter out soft-deleted records
            oBindingParams.filters = oBindingParams.filters || [];
            oBindingParams.filters.push(new Filter("isActive", FilterOperator.EQ, true));

            // Apply status filter
            if (this._sActiveFilterKey && this._sActiveFilterKey !== FILTER_KEY.ALL) {
                const sODataStatus = FILTER_KEY_TO_STATUS[this._sActiveFilterKey];
                if (sODataStatus) {
                    oBindingParams.filters.push(
                        new Filter("announcementStatus", FilterOperator.EQ, sODataStatus)
                    );
                }
            }

            // Apply title search filter
            if (this._sSearchQuery && this._sSearchQuery.trim() !== "") {
                oBindingParams.filters.push(
                    new Filter("title", FilterOperator.Contains, this._sSearchQuery.trim())
                );
            }

            // After data arrives, refresh the filter-button counts
            oBindingParams.events = oBindingParams.events || {};
            const fnOriginalDataReceived = oBindingParams.events.dataReceived;
            oBindingParams.events.dataReceived = (oEvt) => {
                if (fnOriginalDataReceived) { fnOriginalDataReceived(oEvt); }
                this._updateStatusCounts();
            };
        },

        onStatusFilterPress: function (oEvent) {
            const sFilterKey = oEvent.getSource().data("filterKey");
            if (!sFilterKey) { return; }

            this._sActiveFilterKey = sFilterKey;
            this.getOwnerComponent().getModel("announcementCountModel")
                .setProperty("/activeFilter", sFilterKey);

            this.refreshSmartTable();
        },

        _updateStatusCounts: function () {
            const oModel = this.getOwnerComponent().getModel("announcementModel");
            if (!oModel) { return; }

            oModel.read("/Announcements", {
                filters: [new Filter("isActive", FilterOperator.EQ, true)],
                urlParameters: { $select: "announcementStatus" },
                success: (oData) => {
                    const aResults = oData?.results || [];
                    let nAll = aResults.length, nPublished = 0, nScheduled = 0, nExpired = 0, nInvalid = 0;

                    aResults.forEach(({ announcementStatus }) => {
                        if (announcementStatus === STATUS.PUBLISHED) { nPublished++; }
                        else if (announcementStatus === STATUS.TO_BE_PUBLISHED) { nScheduled++; }
                        else if (announcementStatus === STATUS.EXPIRED) { nExpired++; }
                        else if (announcementStatus === STATUS.INVALID) { nInvalid++; }
                    });

                    this.getOwnerComponent().getModel("announcementCountModel").setData({
                        activeFilter: this._sActiveFilterKey,
                        allCount: nAll,
                        publishedCount: nPublished,
                        scheduledCount: nScheduled,
                        expiredCount: nExpired,
                        invalidCount: nInvalid
                    });
                },
                error: (oErr) => { console.error("Failed to fetch announcement counts:", oErr); }
            });
        },

        refreshSmartTable: function () {
            const oSmartTable = this.byId("idAnnouncementsSmartTbl");
            const oModel = this.getOwnerComponent().getModel("announcementModel");

            if (oModel) { oModel.refresh(true); }

            if (oSmartTable) {
                const oBinding = oSmartTable.getTable()?.getBinding("items");
                if (oBinding) { oBinding.refresh(true); }
                oSmartTable.rebindTable();
            }

            this._updateStatusCounts();
        },

        onRefreshPress: function () {
            this.refreshSmartTable();
            MessageToast.show("Table refreshed");
        },


        onSearchAnnouncement: function (oEvent) {
            // Fired when user presses Enter or the search icon
            this._sSearchQuery = oEvent.getParameter("query") || "";
            this.refreshSmartTable();
        },

        onSearchLiveChange: function (oEvent) {
            // Fired on every keystroke — only trigger when field is cleared
            const sValue = oEvent.getParameter("newValue") || "";
            if (sValue === "") {
                this._sSearchQuery = "";
                this.refreshSmartTable();
            }
        },

        onCreateSidebarPress: function () {
            this._router.navTo("CreateSidebarAnnouncement");
        },

        onCreateBannerPress: function () {
            this._router.navTo("CreateBannerAnnouncement");
        },

        onEditPress: function (oEvent) {
            const oData = this._getRowData(oEvent);
            if (!oData) { return; }

            const { announcementId, announcementType } = oData;
            if (!announcementId) { return MessageBox.error("Unable to get announcement ID."); }

            this._navigateToForm(announcementType, announcementId, false);
        },

        onDuplicatePress: function (oEvent) {
            const oData = this._getRowData(oEvent);
            if (!oData) { return; }

            const { announcementId, announcementType } = oData;
            if (!announcementId) { return MessageBox.error("Unable to get announcement ID."); }

            this._navigateToForm(announcementType, announcementId, true);
        },

        _getRowData: function (oEvent) {
            const oBindingCtx = oEvent.getSource().getParent().getParent().getBindingContext();
            if (!oBindingCtx) {
                MessageBox.error("Unable to get announcement data. Please refresh and try again.");
                return null;
            }
            return oBindingCtx.getObject();
        },

        _navigateToForm: function (sAnnouncementType, sAnnouncementId, bDuplicate) {
            const oParams = { announcementId: sAnnouncementId };
            if (bDuplicate) { oParams.duplicate = "true"; }

            if (sAnnouncementType === "Banner") {
                this._router.navTo("CreateBannerAnnouncement", oParams);
            } else if (sAnnouncementType === "Sidebar" || sAnnouncementType === "Sidebar (Popup)") {
                this._router.navTo("CreateSidebarAnnouncement", oParams);
            } else {
                MessageBox.error(`Unknown announcement type: ${sAnnouncementType}. Cannot proceed.`);
            }
        },

        onMarkInvalidPress: function (oEvent) {
            const oData = this._getRowData(oEvent);
            if (!oData) { return; }

            const { announcementId, title } = oData;
            if (!announcementId) {
                return MessageBox.error("Unable to mark as invalid: Announcement ID not found.");
            }

            MessageBox.confirm(`Are you sure you want to mark '${title}' as Invalid?`, {
                title: "Confirm Mark as Invalid",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.NO,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        this._markAnnouncementInvalid(announcementId, title);
                    }
                }
            });
        },

        _markAnnouncementInvalid: function (sAnnouncementId, sTitle) {
            const oBusy = new sap.m.BusyDialog({ text: "Marking announcement as invalid..." });
            oBusy.open();

            this._getCSRFToken()
                .then((csrfToken) => {
                    $.ajax({
                        url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${sAnnouncementId}')`,
                        method: "PATCH",
                        contentType: "application/json",
                        dataType: "json",
                        headers: { "X-CSRF-Token": csrfToken },
                        data: JSON.stringify({ announcementStatus: "INVALID" }),
                        success: () => {
                            oBusy.close();
                            MessageToast.show(`Announcement '${sTitle}' marked as invalid successfully!`, {
                                duration: 4000, width: "25rem",
                                my: "center center", at: "center center", of: window,
                                autoClose: true, animationDuration: 500
                            });
                            setTimeout(() => { this.refreshSmartTable(); }, 500);
                        },
                        error: (xhr) => {
                            oBusy.close();
                            MessageBox.error(
                                xhr.responseJSON?.error?.message || "Failed to mark as invalid. Please try again."
                            );
                        }
                    });
                })
                .catch(() => {
                    oBusy.close();
                    MessageBox.error("Failed to initialize request. Please try again.");
                });
        },

        onDeletePress: function (oEvent) {
            const oData = this._getRowData(oEvent);
            if (!oData) { return; }

            const { announcementId, title } = oData;
            if (!announcementId) {
                return MessageBox.error("Unable to delete: Announcement ID not found.");
            }

            MessageBox.confirm(`Are you sure you want to delete '${title}'?`, {
                title: "Confirm Delete",
                actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                emphasizedAction: MessageBox.Action.NO,
                onClose: (oAction) => {
                    if (oAction === MessageBox.Action.YES) {
                        this._deleteItem(announcementId, title);
                    }
                }
            });
        },

        _deleteItem: function (sAnnouncementId, sTitle) {
            const oBusy = new sap.m.BusyDialog({ text: "Deleting announcement..." });
            oBusy.open();

            this._getCSRFToken()
                .then((csrfToken) => {
                    $.ajax({
                        url: `/JnJ_Workzone_Portal_Destination_Node/odata/v2/announcement/Announcements('${sAnnouncementId}')`,
                        method: "DELETE",
                        contentType: "application/json",
                        dataType: "json",
                        headers: { "X-CSRF-Token": csrfToken },
                        success: () => {
                            oBusy.close();
                            MessageToast.show(`Announcement '${sTitle}' deleted successfully!`, {
                                duration: 4000, width: "25rem",
                                my: "center center", at: "center center", of: window,
                                autoClose: true, animationDuration: 500
                            });
                            setTimeout(() => { this.refreshSmartTable(); }, 500);
                        },
                        error: (xhr) => {
                            oBusy.close();
                            MessageBox.error(
                                xhr.responseJSON?.error?.message || "Failed to delete announcement. Please try again."
                            );
                        }
                    });
                })
                .catch(() => {
                    oBusy.close();
                    MessageBox.error("Failed to initialize request. Please try again.");
                });
        }
    });
});