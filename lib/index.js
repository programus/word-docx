const { isBrowser } = require('browser-or-node')

if (!isBrowser) {
  const { JSDOM } = require('jsdom')
  global.window = new JSDOM().window
  global.document = window.document
  global.DOMParser = window.DOMParser
}

const classes = require('./classes')

window.customElements.define('word-docx', classes.WordDocx)

// for browser
window.$docx = classes

module.exports = classes
