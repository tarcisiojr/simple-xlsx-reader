const unzipper = require('unzipper');
const pipe = require('multipipe')
const { Transform, Writable } = require('stream');
const { chainPromises } = require('./utils');
const xmlParser = require('./xml-stream-parser');

const loadSheet = (stream, sharedStrings) => (entry) => {
    return new Promise((resolve, reject) => {
        entry
            .pipe(xmlParser())
            .pipe(new Transform({
                objectMode: true,
                transform(chunk, _, callback) {
                    if (chunk.name === 'row') {
                        stream.push(parseRow(parseCol(sharedStrings))(chunk))
                    }
                    callback()
                }
            }))
            .on('error', reject)
            .on('finish', resolve);
    })
}

const loadSheets = (stream, sharedStrings, entries) => {
    const loadSheetsWithStream = loadSheet(stream, sharedStrings);

    return chainPromises(entries.map(entry => loadSheetsWithStream.bind(null, entry)));
}

const parseCol = (sharedStrings) => (col) => {
    const value = col.children.filter(e => e.name == 'v')[0].text;

    if (col.attrs['t'] && col.attrs.t === 's') {
        return sharedStrings[parseInt(value)].children[0].text;
    }

    return value;
}

const parseRow = (parseCol) => (row) => {
    return row.children.map(parseCol)
}

const xlsxParser = () => {
    const sharedStrings = []
    const sheetEntries = []

    const loadSharedStrings = (entry, cb) => {
        entry
            .pipe(xmlParser())
            .pipe(new Writable({
                objectMode: true,
                write(chunk, _, callback) {
                    if (chunk.name === 'si') {
                        sharedStrings.push(chunk)
                    }
                    callback();
                }
            }))
            .on('finish', cb);
    }

    const xlsxStream = new Transform({
        objectMode: true,
        transform(entry, _, cb) {
            const fileName = entry.path;
            const processSheets = () => {
                if (sharedStrings.length) {
                    loadSheets(this, sharedStrings, sheetEntries).then(() => {
                        sheetEntries.splice(0);
                        cb();
                    }).catch(err => {
                        console.error(err);
                        cb(err);
                    })
                } else {
                    cb();
                }
            };

            if (fileName === "xl/sharedStrings.xml") {
                loadSharedStrings(entry, processSheets);

            } else if (fileName.match(/xl\/worksheets\/sheet/g)) {
                sheetEntries.push(entry);
                processSheets();

            } else {
                entry.autodrain();
                processSheets();
            }
        }
    });

    return pipe(unzipper.Parse(), xlsxStream);
}

module.exports = xlsxParser