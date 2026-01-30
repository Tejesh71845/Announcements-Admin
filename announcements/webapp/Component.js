sap.ui.define([
    "sap/ui/core/UIComponent",
    "com/incture/announcements/model/models",
    "sap/m/MessageBox",
    "sap/ui/core/Fragment"
], (UIComponent, models, MessageBox, Fragment) => {
    "use strict";

    return UIComponent.extend("com.incture.announcements.Component", {
        metadata: {
            manifest: "json",
            interfaces: [
                "sap.ui.core.IAsyncContentCreation"
            ]
        },

        init() {
            // call the base component's init function
            UIComponent.prototype.init.apply(this, arguments);

            // Initialize session expiration handler
            this._initSessionExpirationHandler();

            // set the device model
            this.setModel(models.createDeviceModel(), "device");

            // enable routing
            this.getRouter().initialize();
        },

        /**
         * Initialize global session expiration handler
         * @private
         */
        _initSessionExpirationHandler() {
            const oAnnouncementModel = this.getModel("announcementModel");
            
            // Attach to OData model request failed event
            if (oAnnouncementModel) {
                oAnnouncementModel.attachRequestFailed(this._handleRequestFailed, this);
            }

            // Also handle jQuery AJAX errors globally for non-OData calls
            $(document).ajaxError((event, jqXHR, ajaxSettings, thrownError) => {
                if (jqXHR.status === 401) {
                    this._showSessionExpiredDialog();
                }
            });
        },

        /**
         * Handle OData request failures
         * @param {object} oEvent - Request failed event
         * @private
         */
        _handleRequestFailed(oEvent) {
            const oResponse = oEvent.getParameter("response");
            
            if (oResponse && oResponse.statusCode === "401") {
                // Prevent default error handling
                oEvent.preventDefault();
                
                this._showSessionExpiredDialog();
            }
        },

        /**
         * Show session expired dialog
         * @private
         */
        _showSessionExpiredDialog() {
            // Prevent multiple dialogs
            if (this._sessionExpiredDialogOpen) {
                return;
            }

            this._sessionExpiredDialogOpen = true;

            if (!this._oSessionExpiredDialog) {
                Fragment.load({
                    name: "com.incture.announcements.fragments.SessionExpired",
                    controller: this
                }).then((oDialog) => {
                    this._oSessionExpiredDialog = oDialog;
                    this._oSessionExpiredDialog.open();
                }).catch((error) => {
                    console.error("Failed to load session expired dialog:", error);
                    this._sessionExpiredDialogOpen = false;
                    // Fallback to simple message
                    this._showFallbackMessage();
                });
            } else {
                this._oSessionExpiredDialog.open();
            }
        },

        /**
         * Handle refresh button press in session expired dialog
         */
        onSessionExpiredRefresh() {
            // Close dialog if open
            if (this._oSessionExpiredDialog) {
                this._oSessionExpiredDialog.close();
            }

            // Clear the flag
            this._sessionExpiredDialogOpen = false;

            // Reload the application
            window.location.reload();
        },

        /**
         * Fallback message if dialog fails to load
         * @private
         */
        _showFallbackMessage() {
            MessageBox.warning(
                "Your session has expired. Please refresh the page to continue.",
                {
                    title: "Session Expired",
                    actions: [MessageBox.Action.OK],
                    emphasizedAction: MessageBox.Action.OK,
                    onClose: () => {
                        window.location.reload();
                    }
                }
            );
        },

        /**
         * Cleanup on component destroy
         */
        exit() {
            if (this._oSessionExpiredDialog) {
                this._oSessionExpiredDialog.destroy();
                this._oSessionExpiredDialog = null;
            }

            const oAnnouncementModel = this.getModel("announcementModel");
            if (oAnnouncementModel) {
                oAnnouncementModel.detachRequestFailed(this._handleRequestFailed, this);
            }
        }
    });
});