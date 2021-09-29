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
const getColumnArray = (columnNumber) => {

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
 *@module excelcolumncharacters
 */
module.exports = {
    getColumnName,
    getColumnArray
};