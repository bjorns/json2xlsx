const optparse = require('optparse');
const xlsx = require('xlsx');
const fs = require('fs');

const SWITCHES = [
    ['-h', '--help', 'Shows help sections'],
    ['-v', '--verbose', 'Output verbose logging'],
    ['-k', '--key COL', 'Set a key column to be placed in column A'],
    ['-o', '--out FILENAME', 'Write to filename (default: out.xlsx)']
];

let verbose = false;

const log = msg => {
    if (verbose)
        console.error(msg);
}

const readJson = filename => {
    log(`Loading file ${filename}`);
    return JSON.parse(fs.readFileSync(filename));
};

const headers = data => {
    const firstItem = data[0];
    return Object.keys(firstItem);
};

const writeToExcel = (filename, data) => {
    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(data, {header: headers(data), cellDates: true});
    let sheetName = 'Data';
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
    const buffer = xlsx.write(workbook, {type: 'buffer', bookType: 'xlsx'});
    fs.writeFileSync(filename, buffer);
};

const main = (args) => {
    const options = new optparse.OptionParser(SWITCHES);
    let keyColumn = null;
    let inputFilename = null;
    let outputFilename = 'out.xlsx';

    options.on('help', () => {
        console.error(optparser.toString());
        process.exit(1);
    });

    options.on('key', function(param, value) {
        keyColumn = value;
    });

    options.on('verbose', () => {
        verbose = true;
    });
    options.on(2, function (opt) {
        inputFilename = opt;
    });

    options.parse(args);

    const data = readJson(inputFilename);

    if (!Array.isArray(data)) {
        throw new Error('Data is not array');
    }

    writeToExcel(outputFilename, data);
};


exports.main = main;