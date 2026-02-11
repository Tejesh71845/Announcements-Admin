sap.ui.define([], function () {
    "use strict";

    return {

        /**
         * Formats the announcement status to appropriate ObjectStatus state
         * @param {string} sStatus - The announcement status
         * @returns {string} The state for ObjectStatus
         */
        formatStatusState: function (sStatus) {
            switch (sStatus) {
                case "PUBLISHED":
                    return "Success";
                case "TO_BE_PUBLISHED":
                    return "Warning";
                case "DRAFT":
                    return "Information";
                default:
                    return "None";
            }
        },

        /**
         * Formats the announcement status text for display
         * @param {string} sStatus - The announcement status
         * @returns {string} The formatted status text
         */
        formatStatusText: function (sStatus) {
            switch (sStatus) {
                case "PUBLISHED":
                    return "Published";
                case "TO_BE_PUBLISHED":
                    return "To Be Published";
                case "DRAFT":
                    return "Draft";
                default:
                    return sStatus || "N/A";
            }
        },
        /**
         * Formats a date to display format (dd/MM/yyyy)
         * @param {Date|string} oDate - Date object or ISO string
         * @returns {string} Formatted date string
         */
        formatDateToDisplay: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            return `${day}/${month}/${year}`;
        },

        /**
         * Formats a date to value format for DatePicker (yyyy-MM-dd)
         * @param {Date} oDate - Date object
         * @returns {string} Formatted date string
         */
        formatDateToValue: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        },

        formatDateToDDMMYYYY: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();

            return `${day}/${month}/${year}`;
        },

        /**
         * Formats a date/time to US locale string
         * @param {Date|string} oDate - Date object or ISO string
         * @returns {string} Formatted date-time string
         */
        formatUSDateTime: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            return date.toLocaleString("en-US", {
                year: "numeric",
                month: "2-digit",
                day: "2-digit",
                hour: "2-digit",
                minute: "2-digit",
                hour12: true
            });
        },

        /**
         * Formats a date to US date format without time (MM/DD/YYYY)
         * @param {Date|string} oDate - Date object or ISO string
         * @returns {string} Formatted date string
         */
        formatDateOnly: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            const year = date.getFullYear();

            return `${month}/${day}/${year}`;
        },

        formatCategoryNames: function (aToTypes) {

            console.log("Formatter received:", aToTypes);
            if (!aToTypes || !Array.isArray(aToTypes)) {
                return "";
            }

            // Extract names directly from the type objects
            const aCategoryNames = aToTypes
                .map(oItem => oItem.type && oItem.type.name)  // Get the name directly
                .filter(Boolean);  // Remove any undefined/null values

            return aCategoryNames.join(", ");
        }

    };
});