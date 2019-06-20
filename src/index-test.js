const fs = require('fs');
const unzipper = require('unzipper');
const { Writable } = require('stream');
const xlsxParser = require('./xlsx-stream-parser');
var XLSX = require('xlsx');

// const fileContents = fs.createReadStream('/home/tarcisio/Temp/Excel.xlsx');
// const fileContents = fs.createReadStream('/home/tarcisio/Temp/Pedido_de_compras_sem_macros_sem_multiplas_abas_ORIGINAL.xlsx');

function impl() {
  // const fileContents = fs.createReadStream('/home/tarcisio/Temp/big.xlsx');
  const fileContents = fs.createReadStream('/home/tarcisio/Temp/Pedido_de_compras_sem_macros_sem_multiplas_abas_ORIGINAL.xlsx');


  let total = 0;

  fileContents
    // .pipe(unzipper.Parse())
    .pipe(xlsxParser())
    .pipe(new Writable({
      objectMode: true,
      // highWaterMark: 5,
      write(chunk, _, callback) {
        // console.log(chunk);
        total++;
        // setTimeout(() => {
          console.log(chunk);
          callback()
        // }, 200)
      }
    }))
    .on('error', (err) => {
      console.log('Error', err);
    })
    .on('finish', () => {
      console.log('total:', total);
    })
    .on('end', () => {
      console.log('total:', total);
    })
}

function xlsx() {
  // const fileContents = fs.createReadStream('/home/tarcisio/Temp/big.xlsx');

  // function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/ {
  //   var buffers = [];
  //   stream.on('data', function (data) { buffers.push(data); });
  //   stream.on('end', function () {
  //     var buffer = Buffer.concat(buffers);
  //     var workbook = XLSX.read(buffer, { type: "buffer" });

  //     /* DO SOMETHING WITH workbook IN THE CALLBACK */
  //     cb(workbook);
  //   });
  // }

  // process_RS(fileContents, (wb) => {
  //   // console.log(wb);
  //   const x = XLSX.utils.sheet_to_json(wb);

  //   console.log(x);
  // })

  const workbook = XLSX.readFile('/home/tarcisio/Temp/big.xlsx');

  var sheet_name_list = workbook.SheetNames;
  const total = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {raw: true})
  console.log(total.length)
}

function gerar() {
  const streamWriter = fs.createWriteStream('/home/tarcisio/Temp/big.csv');

  streamWriter.write('a,b,c,d,e\n');
  for (let i = 0; i < 1000000; i++) {
    streamWriter.write(new Array(5).fill(i).join(',') + '\n')
  }
  streamWriter.end();
}

// xlsx();
// impl();
// gerar();