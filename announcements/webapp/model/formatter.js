sap.ui.define([], function () {
    "use strict";

    return {
        /**
         * Format date to US format
         * @param {Date|string} oDate - Date to format
         * @returns {string} Formatted date string
         */
        formatDateOnly: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const options = { day: '2-digit', month: 'short', year: 'numeric' };
            return new Intl.DateTimeFormat('en-GB', options).format(date);
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


        formatDateToValue: function (oDate) {
            if (!oDate) return "";

            const date = oDate instanceof Date ? oDate : new Date(oDate);
            if (isNaN(date.getTime())) return "";

            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');

            return `${year}-${month}-${day}`;
        },

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
         * Format date for display (without time)
         * @param {string} sDate - Date string
         * @returns {string} Formatted date
         */
        formatDisplayDate: function (sDate) {
            if (!sDate) return "";

            try {
                const date = new Date(sDate);
                if (isNaN(date.getTime())) return sDate;

                return date.toLocaleDateString("en-US", {
                    year: "numeric",
                    month: "2-digit",
                    day: "2-digit"
                });
            } catch (error) {
                return sDate;
            }
        },

        /**
         * Format category text (uppercase first letter)
         * @param {string} sCategory - Category text
         * @returns {string} Formatted category
         */
        formatCategory: function (sCategory) {
            if (!sCategory) return "";
            return sCategory.charAt(0).toUpperCase() + sCategory.slice(1);
        },

        /**
         * Truncate text with ellipsis
         * @param {string} sText - Text to truncate
         * @param {number} iMaxLength - Maximum length
         * @returns {string} Truncated text
         */
        truncateText: function (sText, iMaxLength) {
            if (!sText) return "";
            iMaxLength = iMaxLength || 100;

            if (sText.length <= iMaxLength) {
                return sText;
            }
            return sText.substring(0, iMaxLength) + "...";
        }
    };
});