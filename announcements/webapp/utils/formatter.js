sap.ui.define([], function () {
    "use strict";

    /**
     * Status constants — exact OData values as returned by the backend.
     * OData uses UPPER_SNAKE_CASE.  "Scheduled" in the UI maps to TO_BE_PUBLISHED.
     */
    var STATUS = {
        PUBLISHED: "PUBLISHED",
        TO_BE_PUBLISHED: "TO_BE_PUBLISHED",   // displayed as "Scheduled"
        EXPIRED: "EXPIRED",
        INVALID: "INVALID"
    };

    /**
     * Filter key constants — UI-facing keys used in the count model and customData.
     * These are independent of the OData status strings.
     */
    var FILTER_KEY = {
        ALL: "all",
        PUBLISHED: "published",
        SCHEDULED: "scheduled",   // maps to TO_BE_PUBLISHED in OData
        EXPIRED: "expired",
        INVALID: "invalid"
    };

    /**
     * Maps a UI filter key → OData status value.
     * Used in the controller when building OData filters.
     */
    var FILTER_KEY_TO_STATUS = {
        published: STATUS.PUBLISHED,
        scheduled: STATUS.TO_BE_PUBLISHED,
        expired: STATUS.EXPIRED,
        invalid: STATUS.INVALID
    };

    return {

        /* expose mapping so the controller can import it */
        FILTER_KEY_TO_STATUS: FILTER_KEY_TO_STATUS,
        STATUS: STATUS,
        FILTER_KEY: FILTER_KEY,

        /* ========================================
         * STATUS OBJECT-STATUS FORMATTERS
         * ======================================== */

        /**
         * Human-readable label for the ObjectStatus cell.
         * "TO_BE_PUBLISHED" → "Scheduled"
         */
        formatStatusText: function (sStatus) {
            switch (sStatus) {
                case STATUS.PUBLISHED: return "Published";
                case STATUS.TO_BE_PUBLISHED: return "Scheduled";
                case STATUS.EXPIRED: return "Expired";
                case STATUS.INVALID: return "Invalid";
                default: return sStatus || "";
            }
        },

        /**
         * sap.ui.core.ValueState for ObjectStatus.
         */
        formatStatusState: function (sStatus) {
            switch (sStatus) {
                case STATUS.PUBLISHED: return "Success";
                case STATUS.TO_BE_PUBLISHED: return "Warning";
                case STATUS.EXPIRED: return "Error";
                case STATUS.INVALID: return "None";
                default: return "None";
            }
        },

        /**
         * Icon URI for ObjectStatus.
         */
        formatStatusIcon: function (sStatus) {
            switch (sStatus) {
                case STATUS.PUBLISHED: return "sap-icon://sys-enter-2";
                case STATUS.TO_BE_PUBLISHED: return "sap-icon://pending";
                case STATUS.EXPIRED: return "sap-icon://error";
                case STATUS.INVALID: return "sap-icon://sys-cancel-2";
                default: return "";
            }
        },

        /* ========================================
         * STATUS FILTER BUTTON FORMATTERS
         * ======================================== */

        /**
         * Combines an i18n label with a dynamic count for filter button text.
         * Binding example:
         *   parts: [{path:'i18n>filterBtnPublished'}, {path:'announcementCountModel>/publishedCount'}]
         *   formatter: '.formatter.formatFilterBtnLabel'
         * Output: "Published (5)"
         */
        formatFilterBtnLabel: function (sLabel, nCount) {
            var sText = sLabel || "";
            var nSafeCount = (nCount !== undefined && nCount !== null) ? nCount : 0;
            return sText + " (" + nSafeCount + ")";
        },

        /**
         * A filter button is enabled only when its count > 0.
         */
        formatFilterBtnEnabled: function (nCount) {
            return (nCount !== undefined && nCount !== null && nCount > 0);
        },

        /* One type-formatter per filter key so each button binding stays a single path */
        formatFilterBtnTypeAll: function (sActiveFilter) {
            return sActiveFilter === FILTER_KEY.ALL ? "Emphasized" : "Default";
        },
        formatFilterBtnTypePublished: function (sActiveFilter) {
            return sActiveFilter === FILTER_KEY.PUBLISHED ? "Emphasized" : "Default";
        },
        formatFilterBtnTypeScheduled: function (sActiveFilter) {
            return sActiveFilter === FILTER_KEY.SCHEDULED ? "Emphasized" : "Default";
        },
        formatFilterBtnTypeExpired: function (sActiveFilter) {
            return sActiveFilter === FILTER_KEY.EXPIRED ? "Emphasized" : "Default";
        },
        formatFilterBtnTypeInvalid: function (sActiveFilter) {
            return sActiveFilter === FILTER_KEY.INVALID ? "Emphasized" : "Default";
        },

        /* ========================================
         * ACTION ICON VISIBILITY FORMATTERS
         *
         * Rules:
         *   Edit          → PUBLISHED + TO_BE_PUBLISHED (Scheduled)
         *   Duplicate     → all statuses  (always visible — no formatter needed in view)
         *   Mark Invalid  → PUBLISHED only
         *   Delete        → TO_BE_PUBLISHED (Scheduled) only
         *
         * For rows where an action is NOT available we show a DISABLED placeholder
         * button (same icon, enabled=false) so the Actions column width stays stable.
         * ======================================== */

        /** Edit — active */
        formatEditBtnVisibility: function (sStatus) {
            return sStatus === STATUS.PUBLISHED || sStatus === STATUS.TO_BE_PUBLISHED;
        },

        /** Edit — disabled placeholder (Expired / Invalid rows) */
        formatEditBtnDisabledVisibility: function (sStatus) {
            return sStatus !== STATUS.PUBLISHED && sStatus !== STATUS.TO_BE_PUBLISHED;
        },

        /** Mark As Invalid — active (PUBLISHED only) */
        formatMarkInvalidBtnVisibility: function (sStatus) {
            return sStatus === STATUS.PUBLISHED;
        },

        /** Mark As Invalid — disabled placeholder (all except PUBLISHED) */
        formatMarkInvalidBtnDisabledVisibility: function (sStatus) {
            return sStatus !== STATUS.PUBLISHED;
        },

        /** Delete — active (TO_BE_PUBLISHED / Scheduled only) */
        formatDeleteBtnVisibility: function (sStatus) {
            return sStatus === STATUS.TO_BE_PUBLISHED;
        },

        /** Delete — disabled placeholder (all except TO_BE_PUBLISHED) */
        formatDeleteBtnDisabledVisibility: function (sStatus) {
            return sStatus !== STATUS.TO_BE_PUBLISHED;
        },

        /* ========================================
         * DATE FORMATTERS
         * ======================================== */

        formatDateOnly: function (oDate) {
            if (!oDate) { return ""; }
            var dDate = (oDate instanceof Date) ? oDate : new Date(oDate);
            if (isNaN(dDate.getTime())) { return ""; }
            var sMonth = String(dDate.getMonth() + 1).padStart(2, "0");
            var sDay = String(dDate.getDate()).padStart(2, "0");
            var sYear = dDate.getFullYear();
            return sMonth + "/" + sDay + "/" + sYear;
        },

        formatUSDateTime: function (oDate) {
            if (!oDate) { return ""; }
            var dDate = (oDate instanceof Date) ? oDate : new Date(oDate);
            if (isNaN(dDate.getTime())) { return ""; }
            return dDate.toLocaleString("en-US", {
                year: "numeric", month: "2-digit", day: "2-digit",
                hour: "2-digit", minute: "2-digit", hour12: true
            });
        },

        formatDateToDisplay: function (oDate) {
            return this.formatDateOnly(oDate);
        },

        formatDateToValue: function (oDate) {
            if (!oDate) { return ""; }
            var dDate = (oDate instanceof Date) ? oDate : new Date(oDate);
            if (isNaN(dDate.getTime())) { return ""; }
            var sYear = dDate.getFullYear();
            var sMonth = String(dDate.getMonth() + 1).padStart(2, "0");
            var sDay = String(dDate.getDate()).padStart(2, "0");
            return sYear + "-" + sMonth + "-" + sDay;
        },

        formatDateToDDMMYYYY: function (oDate) {
            if (!oDate) { return ""; }
            var dDate = (oDate instanceof Date) ? oDate : new Date(oDate);
            if (isNaN(dDate.getTime())) { return ""; }
            var sDay = String(dDate.getDate()).padStart(2, "0");
            var sMonth = String(dDate.getMonth() + 1).padStart(2, "0");
            var sYear = dDate.getFullYear();
            return sDay + "/" + sMonth + "/" + sYear;
        },

        formatCategoryArray: function (vCategory) {
            if (!vCategory) { return ""; }
            if (Array.isArray(vCategory)) {
                return vCategory.join(", ");
            }
            return typeof vCategory === "string" ? vCategory : "";
        },
    };
});