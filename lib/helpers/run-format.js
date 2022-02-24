const { object2CssText } = require('./utils')

/**
 * http://officeopenxml.com/WPtextFormatting.php
 */
const convertors = {
  b(style, obj) {
    style['font-weight'] = obj.val === 'false' ? 'normal' : 'bold'
  },

  i(style, obj) {
    style['font-style'] = obj.val === 'false' ? 'normal' : 'italic'
  },

  caps(style, obj) {
    style['text-transform'] = obj.val === 'false' ? 'none' : 'uppercase'
  },

  color(style, obj) {
    const color = obj.val
    style.color = `#${color}`
  },

  dstrike(style, obj) {
    style['text-decoration'] = 'line-through double'
  },

  smallCaps(style, obj) {
    style['font-variant'] = obj.val === 'false' ? 'normal' : 'small-caps'
  },

  strike(style, obj) {
    style['text-decoration'] = 'line-through'
  },

  sz(style, obj) {
    const value = obj.val
    style['font-size'] = `${value / 2}pt`
  },

  szCs(style, obj) {
    this.sz(style, obj)
  },

  u(style, obj) {
    const value = obj.val
    if (value !== 'none') {
      const styleMap = {
        dash: 'dashed',
        dashedHeavy: 'dashed',
        dotted: 'dotted',
        dottedHeavy: 'dotted',
        double: 'double',
        wave: 'wavy',
        wavyHeavy: 'wavy',
      }
      style['text-decoration-style'] = styleMap[value] || 'solid'
      style['text-decoration-thickness'] = value.endsWith('Heavy') || value === 'thick' ? '2pt' : 'auto'
      style['text-decoration-line'] = 'underline'
      const color = obj.color
      if (color) {
        style['text-decoration-color'] = `#${color}`
      }
    }
  },

  bCs(style, obj) {
    this.b(style, obj)
  },

  bdo(style, obj) {
    style['direction'] = obj.val
    style['unicode-bidi'] = 'bidi-override'
  },

  shd(style, obj) {
    // TODO: w:val & w:color implementation
    style['background-color'] = `#${obj.fill}`
  },
}

function elements2Json(elements) {
  return Object.fromEntries(Array.from(elements).map(e => [
    e.tagName.substring('w:'.length),
    Object.fromEntries(Array.from(e.attributes).map(attr => [
      attr.name.substring('w:'.length),
      attr.value,
    ]))
  ]))
}

class RunFormat {
  constructor(params, important) {
    this._formatJson = params.constructor.name === 'Object' ? params : elements2Json(params)
    this.style = this._generateCssText()
    this._important = important
  }

  _generateCssText() {
    const style = {}
    Object.entries(this._formatJson).forEach(e => {
      const convertor = convertors[e[0]]
      if (convertor) {
        convertor.bind(convertors)(style, e[1])
      }
    })
    return style
  }

  get formatJson() {
    return this._formatJson
  }

  get cssText() {
    return object2CssText(this.style, this._important)
  }
}

module.exports = RunFormat