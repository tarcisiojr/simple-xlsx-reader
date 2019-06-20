const unzipper = require('unzipper');
const pipe = require('multipipe')
const { Transform, Writable, Duplex } = require('stream');
const { chainPromises } = require('./utils');
const xmlParser = require('simple-xml-reader');

const chainPromises = (promises) => {
    return promises.reduce((prev, task) => {
        return prev.then(() => task())
    }, Promise.resolve())
}

const emptyChildrenOfSheetData = (elem) => {
    if (elem.name === 'row') {
        elem.parent.children = [];
    }

    return elem;
}

const loadSheet = (push) => (entry) => {
    return new Promise((resolve, reject) => {
        entry
            .pipe(xmlParser({ mapElement: emptyChildrenOfSheetData }))
            .pipe(new Writable({
                objectMode: true,
                write(chunk, _, callback) {
                    if (chunk.name === 'row') {
                        push(chunk, callback);
                    } else {
                        callback();
                    }
                }
            }))
            .on('error', reject)
            .on('finish', () => {
                push(null);
                resolve();
            });
    })
}

const loadSheets = (push, entries) => {
    const loadSheetsWithConfig = loadSheet(push);

    return chainPromises(entries.map(entry => loadSheetsWithConfig.bind(null, entry)));
}

const parseCol = (sharedStrings, styles, date1904) => (col) => {
    // console.log(col);
    const arrValues = col.children.filter(e => e.name == 'v');
    const value = arrValues.length ? arrValues[0].text : null;

    if (value === null) {
        return '';
    }

    if (col.attrs.t === 's') {
        return sharedStrings[parseInt(value)].children[0].text;

    } else if (col.attrs.s !== '0' 
        && col.attrs.t === 'n'
        && isDateFmt(styles[parseInt(col.attrs.s)].attrs.formatCode)) {
        return excelToDate(value, date1904)
    }

    return value;
}

const parseRow = (parseCol) => (row) => {
    return row.children.map(parseCol)
}

const excelToDate = (v, date1904) => {
    const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
    return new Date(millisecondSinceEpoch);
}

const isDateFmt = (fmt) => {
    if (!fmt) {
        return false;
    }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/\[[^\]]*]/g, '');
    fmt = fmt.replace(/"[^"]*"/g, '');
    // then check for date formatting chars
    const result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
}

const load = (fn) => (entry, cb) => {
    entry
        .pipe(xmlParser())
        .pipe(new Writable({
            objectMode: true,
            write(chunk, _, callback) {
                fn(chunk);
                callback();
            }
        }))
        .on('finish', cb);
}

const xlsxParser = () => {
    const sharedStrings = [];
    const sheetEntries = [];
    const styles = [];
    let stylesLoaded = false;
    let workbookLoaded = false;
    let date1904 = false;
    let resume = null;

    const loadSharedStrings = load((chunk) => {
        if (chunk.name === 'si') {
            sharedStrings.push(chunk)
        }
    });

    const loadStyles = load((chunk) => {
        if (chunk.name === 'numFmt') {
            styles.push(chunk)
        }
    });

    const loadWorkbook = load((chunk) => {
        if (chunk.name === 'workbookPr') {
            date1904 = chunk.attrs['date1904'] 
                && (chunk.attrs['date1904'] === 'true' || chunk.attrs['date1904'] === true)
        }
    });

    const canLoadSheets = () => sharedStrings.length && stylesLoaded && workbookLoaded;

    const processSheets = (stream, next) => {
        if (!canLoadSheets()) {
            return next();
        }

        const push = (chunk, callback) => {
            const pause = !stream.push(chunk ? parseRow(parseCol(sharedStrings, styles, date1904))(chunk) : null);
            
            if (pause) {
                resume = callback;
            } else {
                callback && callback();
            }
        }

        loadSheets(push, sheetEntries)
            .then(() => {
                sheetEntries.splice(0);
                next();
            }).catch(err => {
                console.error(err);
                next(err);
            });
    };
    
    const xlsxStream = new Duplex({
        objectMode: true,
        read(size) {
            if (resume) {
                let fn = resume;
                resume = null;
                fn();
            }
        },
        write(entry, _, cb) {
            const fileName = entry.path;

            if (fileName === "xl/sharedStrings.xml") {
                loadSharedStrings(entry, processSheets.bind(null, this, cb));

            } else if (fileName.match(/xl\/worksheets\/sheet/g)) {
                sheetEntries.push(entry);
                processSheets(this, cb);

            } else if (fileName === 'xl/styles.xml') {
                loadStyles(entry, () => {
                    stylesLoaded = true;
                    processSheets(this, cb);
                });

            } else if (fileName === 'xl/workbook.xml') {
                loadWorkbook(entry, () => {
                    workbookLoaded = true;
                    processSheets(this, cb);
                });

            } else {
                entry.autodrain();
                processSheets(this, cb);
            }
        }
    });

    return pipe(unzipper.Parse(), xlsxStream);
}

module.exports = xlsxParser