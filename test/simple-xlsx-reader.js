/* eslint-disable no-undef */
const { xlsxParser, asArray } = require('../src/')
const expect = require('chai').expect
const fs = require('fs')
const path = require('path')
const { Writable } = require('stream')

const write = (fn) => {
  return new Writable({
    objectMode: true,
    write (chunk, _, callback) {
      fn(chunk)
      callback()
    }
  })
}

describe('Testing XLSX Parser', () => {
  it('xlsxParser should be a function', () => {
    expect(xlsxParser).to.be.a('function')
  })

  it('validate nodes of xml', () => {
    const elements = []
    fs.createReadStream(path.resolve(__dirname, './assets/sample1.xlsx'))
      .pipe(xlsxParser(asArray()))
      .pipe(write((chunk) => {
        expect(chunk).to.be.a('array').that.is.not.empty
        elements.push(chunk)
      }))
      .on('finish', () => {
        // expect(elements).has.length(4)
        // expect(elements[0]).to.deep.include({ name: 'child', attrs: { id: '1' } })
        // expect(elements[0].parent).to.deep.include({ name: 'xml' })
        // expect(elements[1]).to.deep.include({ name: 'value', text: '1' })
        // expect(elements[1].parent).to.deep.include({ name: 'child' })
        // expect(elements[2]).to.deep.include({ name: 'child', attrs: { xpto: 'foo' } })
        // expect(elements[2].parent).to.deep.include({ name: 'xml' })
        // expect(elements[3]).to.deep.include({ name: 'xml', parent: null })
      })
  })

  // it('testing map element', () => {
  //   const elements = []
  //   fs.createReadStream(path.resolve(__dirname, './assets/sample1.xml'))
  //     .pipe(xlsxParser({
  //       mapElement: (e) => {
  //         e.xpto = 'foo'
  //         return e
  //       }
  //     }))
  //     .pipe(write((chunk) => {
  //       expect(chunk).to.be.a('object').that.is.not.empty
  //       elements.push(chunk)
  //     }))
  //     .on('finish', () => {
  //       expect(elements).has.length(4)
  //       expect(elements[0]).has.property('xpto', 'foo')
  //     })
  // })
})
