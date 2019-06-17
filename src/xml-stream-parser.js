const expat = require('node-expat');
const { peek } = require('./utils');
const { Transform } = require('stream');

const createElement = (name, attrs, parent) => {
  const element = {
    parent,
    name,
    attrs,
    text: null,
    children: []
  }

  if (parent) {
    parent.children.push(element)
  }

  return element;
}

const xmlParser = () => {
  const root = {};
  const elements = [];
  const parser = new expat.Parser('UTF-8');

  parser.on('startElement', (name, attrs) => {
    const newElement = createElement(name, attrs, peek(elements));

    if (elements.length == 0) {
      Object.assign(root, newElement)
    }

    elements.push(newElement);
  });

  parser.on('text', (text) => {
    peek(elements).text = text
  });

  parser.on('endElement', (name) => {
    stream.push(elements.pop())
  });

  const stream = new Transform({
    objectMode: true,
    write(chunk, _, callback) {
      parser.write(chunk);

      callback();
    }
  });

  return stream;
}

const clearParent = (element) => {
  element.parent = null;
  element.children = element.children.map(e => clearParent(e));
  return element;
}

module.exports = xmlParser
