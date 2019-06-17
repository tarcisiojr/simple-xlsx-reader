const fs = require('fs');
const unzipper = require('unzipper');
const { Writable } = require('stream');
const xlsxParser = require('./xlsx-stream-parser');

const fileContents = fs.createReadStream('/home/tarcisio/Temp/Excel.xlsx');

fileContents
  // .pipe(unzipper.Parse())
  .pipe(xlsxParser())
  .pipe(new Writable({
    objectMode: true,
    write(chunk, _, callback) {
      console.log('-----')
      console.log(chunk)
      callback()
    }
  }))
  .on('error', (err) => {
    console.log('Error', err);
  })