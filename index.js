/**
 * Converts integer value to its corresponding excel column name
 * @param {number} columnNumber - columnNumber as a number
 * @returns {*} columnNumber converted to the particular column name
 */
const getColumnName = (columnNumber) => {
    return new Promise(resolve => {

        // To store result (Excel column name)
        let columnName = [];

        while (columnNumber > 0) {
            // Find remainder
            let rem = columnNumber % 26;

            // If remainder is 0, then a
            // 'Z' must be there in output
            if (rem == 0) {
                columnName.push("Z");
                columnNumber = Math.floor(columnNumber / 26) - 1;
            }
            else // If remainder is non-zero
            {
                columnName.push(String.fromCharCode((rem - 1) + 'A'.charCodeAt(0)));
                columnNumber = Math.floor(columnNumber / 26);
            }
        }

        // Reverse the string and print result
        resolve(columnName.reverse().join(''))
    })
}

/**
 * Creates an array of excel column names till provided column number 
 * @param {number} columnNumber - columnNumber as a number
 * @returns {*} array of excel column names
 */
const getColumnNameArray = (columnNumber) => {

    return new Promise(resolve => {
        // to store all columns
        let columnsNames = [];

        for (let i = 1; i <= columnNumber; i++) {
            let num = i
            // To store result (Excel column name)
            let columnName = [];

            while (num > 0) {
                // Find remainder
                let rem = i % 26;

                // If remainder is 0, then a
                // 'Z' must be there in output
                if (rem == 0) {
                    columnName.push("Z");
                    num = Math.floor(num / 26) - 1;
                }
                else // If remainder is non-zero
                {
                    columnName.push(String.fromCharCode((rem - 1) + 'A'.charCodeAt(0)));
                    num = Math.floor(num / 26);
                }
            }
            // Reverse the string and print result
            columnsNames.push(columnName.reverse().join(''))
        }
        resolve(columnsNames)
    })
}

/**
 * Converts excel column name to its corresponding excel column number
 * @param {number} columnName - columnName as a string
 * @returns {*} columnName converted to the particular column number
 */
const getColumnNumber = (columnName) => {
    return new Promise(resolve => {

        var digits = columnName.toUpperCase().split(''),
            number = 0;

        for (var i = 0; i < digits.length; i++) {
            number += (digits[i].charCodeAt(0) - 64) * Math.pow(26, digits.length - i - 1);
        }

        resolve(number)
    })
}

/**
 * Converts excel column name to its corresponding excel column number
 * @param {number} columnNames - columnName as a array of string 
 * @returns {*} array of excel column numbers
 */
const getColumnNumberArray = (columnNames) => {
    return new Promise(resolve => {
        let excelColumnNumbers = []

        for (let n = 0; n < columnNames.length; n++) {
            var digits = columnNames[n].toUpperCase().split(''),
                number = 0;

            for (var i = 0; i < digits.length; i++) {
                number += (digits[i].charCodeAt(0) - 64) * Math.pow(26, digits.length - i - 1);
            }
            excelColumnNumbers.push(number)
        }
        resolve(excelColumnNumbers)
    })
}

/**
 *@module excelcolumncharacters
 */
module.exports = {
    getColumnName,
    getColumnNameArray,
    getColumnNumber,
    getColumnNumberArray
};
