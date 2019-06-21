
# simple-xlsx-reader
A utility to read xlsx files through NodeJS Stream.

[![Build Status](https://travis-ci.org/tarcisiojr/simple-xlsx-reader.svg?branch=master)](https://travis-ci.org/tarcisiojr/simple-xlsx-reader)

# Instalation

```
npm i simple-xlsx-reader --save
```

# Usage

```javascript
const { Writable } = require('stream')
const fs = require('fs')
const xmlParser = require('simple-xlsx-reader')

const write = (fn) => {
  return new Writable({
    objectMode: true,
    write (chunk, _, callback) {
      fn(chunk)
      callback()
    }
  })
}

fs.createReadStream('path to xlsx file'))
  .pipe(xmlParser())
  .pipe(write((row) => {
    console.log('XLSX Row', row)
  }))
```