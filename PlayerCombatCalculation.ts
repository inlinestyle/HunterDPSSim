import { Row, r0, r1, r2, r3 } from './copypasta';

const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const alphabetLength = alphabet.length;
const startColumn = 'O';
const startColumnIndex = alphabet.indexOf(startColumn);

function identifyColumn(i: number): string {
    // Convert e.g. 0 -> 'O', 1 -> 'P'
    const adjustedI = i + startColumnIndex;
    if (adjustedI < alphabetLength) {
        return alphabet[adjustedI];
    }
    return alphabet[Math.floor(adjustedI / alphabetLength) - 1] + alphabet[adjustedI % alphabetLength];
}

// parseInt(*, 26) would work conceptually except it goes 1-P instead of A-Z (and also starts at 0)
function lookupColumn(column: string): number {
    return Array.from(column).reduce((num, letter) => num * 26 + parseInt(letter, 36) - 9, 0) - startColumnIndex - 1;
}

function identifyIndex(i: number, row: number): string | null {
    switch (i) {
        case row:
            return `[i]`;
        case row - 1:
            return `[i-1]`;
        case row - 2:
            return `[i-2]`;
        default:
            return null;
    }
}

function rstrip(str: string, char: string): string {
    return str.startsWith(char)
        ? str.slice(1)
        : str;
}

function translateMath(str: string): string {
    return str.replace(/ABS/g, 'Math.abs')
        .replace(/sum/g, 'SUM') // "SUM" is referenced as "sum" in some of the formulas
        .replace(/Max/g, 'MAX') // "MAX" is referenced as "Max" in some of the formulas
        .replace(/TRUE/g, 'true')
        .replace(/FALSE/g, 'false')
        .replace(/ROUNDDOWN/g, 'Math.floor')
        .replace(/ROUNDUP/g, 'Math.ceil')
        .replace(/(?<![<>])=/g, '==') // Deliberately using `==` over `===` in hopes that it lets us ignore types
        .replace(/%/g, '')
        .replace(/<>/g, '!=');
}


const localCellReferenceStart = /^\$?([A-Z]{1,3})(\$?)(\d{1,4})/;
const localCellReferenceMiddle = /([^!$:'"A-Z])\$?([A-Z]{1,3})(\$?)(\d{1,4})/g;

const doubleQuotesExternalCellReference = /"([^"]+)"!(\$?[A-Z]{1,3}\$?\d{1,4})/g;
const singleQuotesExternalCellReference = /'([^']+)'!(\$?[A-Z]{1,3}\$?\d{1,4})/g;
const noQuotesExternalCellReference = /(\w+)!(\$?[A-Z]{1,3}\$?\d{1,4})/g;

const startOffset = r0.rowOffset;

function localVarReplacer(str: string, row: number, leadingMatch: boolean) {
    return (match: string, ...matches: string[]): string => {
        let referenceColumn: string, isFixedRow: string, referenceRow: string,
            start = '';
        if (leadingMatch) {
            [start, referenceColumn, isFixedRow, referenceRow] = matches;
        } else {
            [referenceColumn, isFixedRow, referenceRow] = matches;
        }
        const intReferenceRow = parseInt(referenceRow, 10);
        if (isFixedRow || (intReferenceRow < row - 2) || intReferenceRow < startOffset) {
            return `${start}sheet.getRange("${referenceColumn}${referenceRow}").getValue()`;
        }
        const index = identifyIndex(intReferenceRow, row);
        if (index) {
            return `${start}${referenceColumn}${index}`;
        }
        throw new Error(`Unhandled variable expression in "${str}"\nUnhandled expression: "${match}"`);
    }
}

function translateVars(str: string, row: number): string {
    return str.replace(localCellReferenceStart, localVarReplacer(str, row, false))
        .replace(localCellReferenceMiddle, localVarReplacer(str, row, true));
}

function externalVarReplacer(_: unknown, externalSheetName: string, cellReference: string): string {
    return `spreadsheet.getSheetByName('${externalSheetName}').getRange('${cellReference}').getValue()`;
};

function translateExternalVars(str: string): string {
    return str.replace(noQuotesExternalCellReference, externalVarReplacer)
        .replace(doubleQuotesExternalCellReference, externalVarReplacer)
        .replace(singleQuotesExternalCellReference, externalVarReplacer);
}

function translateRanges(str: string, row: number): string {
    return str.replace(/\$?([A-Z]{1,3})(\$?)(\d{1,4}):\$?([A-Z]{1,3})(\$?)(\d{1,4})/g, (match: string, ...matches: string[]): string => {
        const lhs = {
            referenceColumn: matches[0],
            isFixedRow: matches[1],
            referenceRow: matches[2],
        }
        const rhs = {
            referenceColumn: matches[3],
            isFixedRow: matches[4],
            referenceRow: matches[5],
        }
        if (lhs.isFixedRow !== rhs.isFixedRow) {
            throw new Error(`Unhandled range expression in "${str}"\nUnhandled expression: "${match}"`)
        }
        if (lhs.isFixedRow) {
            return `sheet.getRange('${match}').getValues()`;
        }
        // In practice the sheet doesn't contain any multi-row ranges that aren't fixed
        if (lhs.referenceRow === rhs.referenceRow) {
            const index = identifyIndex(parseInt(lhs.referenceRow, 10), row);
            if (index) {
                const intStartingColumn = lookupColumn(lhs.referenceColumn);
                const intEndingColumn = lookupColumn(rhs.referenceColumn);
                const columns = [];
                for (let c = intStartingColumn; c <= intEndingColumn; c++) {
                    columns.push(identifyColumn(c) + index);
                }
                return '[' + columns.join(', ') + ']';
            }
        }
        throw new Error(`Unhandled range expression in "${str}"\nUnhandled expression: "${match}"`)
    });
}

function handleCellEdgeCases(cell: string): string {
    switch (cell) {
        case '':
            return '0';
        case 'Coeffs': // CW31
            return '"Coeffs"';
        case 'Auto': // CW32
            return '"Auto"';
    }
    return cell;
}

function translateRow(row: Row, rowStart: number, iterations: number): string {
    const cells = row.cellsString.split('\t');
    const columns: Array<[string, string]> = [];
    for (const c in cells) {
        const columnString = identifyColumn(parseInt(c, 10));
        if (columnString === 'EP') continue;
        const cell = handleCellEdgeCases(cells[c]);

        let translation = rstrip(cell.trim(), '=');
        translation = translateMath(translation);
        translation = translateRanges(translation, row.rowOffset);
        translation = translateVars(translation, row.rowOffset);
        translation = translateExternalVars(translation);
        columns.push([columnString, translation]);
    }
    const linearizedColumns = linearizeRowColumns(columns);
    return [
        `    for (let i = ${rowStart}; i < ${rowStart + iterations}; i++) {`,
        ...linearizedColumns.map(([columnString, translation]) => `        ${columnString}[i] = ${translation};`),
        logRow(r0, 'i'),
        '    }'
    ].join('\n');
};

function dedupe(array: string[]): string[] { return Array.from(new Set(array)); }

interface SortableCell {
    originalPosition: number;
    columnString: string;
    translation: string;
    dependencies: string[];
}
function linearizeRowColumns(columns: Array<[string, string]>): Array<[string, string]> {
    const cells: SortableCell[] = [];
    for (let i in columns) {
        const [columnString, translation] = columns[i];
        cells.push({
            originalPosition: parseInt(i, 10),
            columnString,
            translation,
            dependencies: dedupe(Array.from(translation.matchAll(/([A-Z]{1,3})\[i]/g)).map(([, d]) => d))
        });
    }

    const inserted = new Set<string>();
    const linearized: SortableCell[] = [];

    // Inefficient but should eventually produce an ordered topological sort
    while (inserted.size !== cells.length) {
        for (const cell of cells) {
            if (inserted.has(cell.columnString)) continue;
            if (cell.dependencies.every(d => inserted.has(d))) {
                inserted.add(cell.columnString);
                linearized.push(cell);
            }
        }
    }

    return linearized.map(({ columnString, translation }) => [columnString, translation]);
}

function initializeColumnArrays({ cellsString }: Row): string {
    const cellCount = cellsString.split('\t').length;
    let str = '';
    for (let c = 0; c < cellCount; c++) {
        str += `    const ${identifyColumn(c)} = [];\n`;
    }
    return str;
}

function logRow({ cellsString }: Row, i: number | string): string {
    const cellCount = cellsString.split('\t').length;
    let str = '    console.log(';
    for (let c = 0; c < cellCount; c++) {
        str += `${identifyColumn(c)}[${i}],`;
    }
    str += ');';
    return str;
}


// FUNCTIONS to write: VLOOKUP, COUNTIF

const header = '/** @OnlyCurrentDoc */';
const if_ = 'function IF(pred, b0, b1) { return pred ? b0 : (b1 || 0); }';
const or_ = 'function OR(...preds) { return preds.some(p => p); }';
const and_ = 'function AND(...preds) { return preds.every(p => p); }';
const max_ = 'function MAX(...nums) { return nums.length > 1 ? Math.max(...nums) : Math.max(...nums[0]); }';
const min_ = 'function MIN(...nums) { return nums.length > 1 ? Math.min(...nums) : Math.min(...nums[0]); }';
const sum_ = 'function SUM(nums) { return nums.reduce((sum, num) => sum + num); }';

// MATCH is only used for exact matches in the sheet
const match_ = 'function MATCH(value, array) { return array.indexOf(value); }';

// INDEX is only used on 2D arrays in the sheet
const index_ = 'function INDEX(matrix, row, column) { return matrix[row][column]; }';

// COUNTIF can technically be used on a range but in this spreadsheet it's only being used on single cells
const countIf_ = 'function COUNTIF(cell, str) { return (cell && cell.includes(str)) ? 1 : 0 }';

const vlookup_ = `function VLOOKUP(value, matrix, returnColumn) {
    for (const row of matrix) {
        if (row[0] === value) {
            return row[returnColumn-1];
        }
    }
}`;

const functionOpen = 'function RunPlayerCombatCalculation() {';
const functionClose = '}';
const spreadsheet = '    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()';
const sheet = '    const sheet = spreadsheet.getSheetByName("Player Combat Calc");';

// The bulk of relevant named ranges are single-cell. Special casing known exceptions.
const setRanges = `
    for (const range of spreadsheet.getNamedRanges()) {
        const name = range.getName();
        if (name === 'CLReferenceTable') {
            this[name] = range.getRange().getValues();
        } else {
            this[name] = range.getRange().getValue();
        }
    }`;

function main() {
    const output: string[] = [
        header,
        and_,
        countIf_,
        if_,
        index_,
        match_,
        max_,
        min_,
        or_,
        sum_,
        vlookup_,
        functionOpen,
        spreadsheet,
        sheet,
        setRanges,
        initializeColumnArrays(r0),
        translateRow(r0, r0.rowOffset, 1),
        translateRow(r1, r1.rowOffset, 1),
        translateRow(r2, r2.rowOffset, 1),
        translateRow(r3, r3.rowOffset, 500),
        functionClose,
    ];
    for (const str of output) { console.log(str); }
}

main();
