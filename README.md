# Excel Column Characters

# Install

```sh
$ npm i excelcolumncharacters
```

# Usage

A simple set of functions to generate character(s) like excel columns names.

## _Excel column name by column number_

Convert a column number to corresponding excel column name.

**_Example_**

```javascript
const { getColumnName } = require("excelcolumncharacters");

const columnName1 = await getColumnName(2); // returns a promise
const columnName2 = await getColumnName(26); // returns a promise
const columnName3 = await getColumnName(51); // returns a promise
const columnName4 = await getColumnName(52); // returns a promise
const columnName5 = await getColumnName(80); // returns a promise
const columnName6 = await getColumnName(676); // returns a promise
const columnName7 = await getColumnName(702); // returns a promise
const columnName8 = await getColumnName(705); // returns a promise

console.log(columnName1);
console.log(columnName2);
console.log(columnName3);
console.log(columnName4);
console.log(columnName5);
console.log(columnName6);
console.log(columnName7);
console.log(columnName8);
```

**_Output_**

```
B
Z
AY
AZ
CB
YZ
ZZ
AAC
```

## _Excel column names in array till given column number_

Generate array of excel column characters till given column number.

**_Example_**

```javascript
const { getColumnArray } = require("excelcolumncharacters");

const columnArray = await getColumnArray(20); // returns a promise

console.log(columnArray);
```

**_Output_**

```
[
 'A', 'B', 'C', 'D', 'E',
 'F', 'G', 'H', 'I', 'J',
 'K', 'L', 'M', 'N', 'O',
 'P', 'Q', 'R', 'S', 'T'
]
```

## _Excel column number from given column name_

Convert a column name to corresponding excel column number.

**_Example_**

```javascript
const { getColumnNumber } = require("excelcolumncharacters");

const columnNumber = await getColumnNumber("AAC"); // returns a promise

console.log(columnNumber);
```

**_Output_**

```
705
```

## _Excel column number from given column name_

Generate array of excel column numbers from provided array of column numbers.

**_Example_**

```javascript
const { getColumnNumberArray } = require("excelcolumncharacters");

const columnNumberArray = await getColumnNumberArray(["a", "aac"]); // returns a promise

console.log(columnNumberArray);
```

**_Output_**

```
[ 1, 705 ]
```
