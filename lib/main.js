const optparse = require('optparse');
const xlsx = require('xlsx');
const fs = require('fs');
const readline = require('readline');


const SWITCHES = [
    ['-h', '--help', 'Shows help sections'],
    ['-v', '--verbose', 'Output verbose logging'],
    ['-k', '--key COL', 'Set a key column to be placed in column A'],
    ['-o', '--out FILENAME', 'Write to filename (default: out.xlsx)'],
    ['-l', '--lines', 'Line mode - File consists of one object per line as opposed to json array'],
];

let verbose = false;

const log = msg => {
    if (verbose)
        console.error(msg);
}

const readJson = (filename, lineMode) => new Promise((resolve, reject) => {
    if (lineMode) {
        const ret = [];
        const reader = readline.createInterface({
            input: fs.createReadStream(filename)
        });
        reader.on('line', line => {
            //console.log(`"${line}"`);
            ret.push(JSON.parse(line));
        });
        reader.on('close', line => {
          resolve(ret);
        });
    } else {
        return JSON.parse(fs.readFile(filename, (err, data) => {
            if (err) {
                reject(err);
            } else {
                resolve(data);
            }
        }));
    }
});

const headers = (data, keyColumn) => {
    const firstItemCopy = Object.assign({}, data[0]);
    let keys = [];
    if (keyColumn) {
        keys.push(keyColumn)
        delete firstItemCopy[keyColumn];
    }
    return keys.concat(Object.keys(firstItemCopy));
};

const writeToExcel = (filename, data, keyColumn) => {
    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(data, {header: headers(data, keyColumn), cellDates: true});
    let sheetName = 'Data';
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
    const buffer = xlsx.write(workbook, {type: 'buffer', bookType: 'xlsx'});
    fs.writeFileSync(filename, buffer);
};

const main = (args) => {
    const options = new optparse.OptionParser(SWITCHES);
    options.banner = 'Usage: json2xlsx [options] <input_file>';
    let keyColumn = null;
    let inputFilename = null;
    let outputFilename = 'out.xlsx';
    let lineMode = false;

    options.on('help', () => {
        console.error(optparser.toString());
        process.exit(1);
    });

    options.on('key', function(param, value) {
        keyColumn = value;
    });
    options.on('out', function(param, value) {
        outputFilename = value;
    });
    options.on('verbose', () => {
        verbose = true;
    });
    options.on('lines', () => {
        lineMode = true;
    });
    options.on(2, function (opt) {
        inputFilename = opt;
    });

    if (args.length < 3) {
        console.error(options.toString());
        process.exit(-1);
    }
    options.parse(args);

    readJson(inputFilename, lineMode).then(data => {
        if (!Array.isArray(data)) {
            throw new Error('Data is not array');
        }
        writeToExcel(outputFilename, data, keyColumn);
    }).catch(err => {
        console.error(`Failed to read file ${inputFilename}`);
        console.error(err);
    });
};


exports.main = main;