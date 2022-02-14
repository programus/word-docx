const convertors = {
  color(style, obj) {
    const color = obj.val
    style.color = `#${color}`
  },

  b(style, obj) {
    style['font-weight'] = obj.val === 'false' ? 'normal' : 'bold'
  },

  bCs(style, obj) {
    this.b(style, obj)
  },

  bdo(style, obj) {
    style['direction'] = obj.val
    style['unicode-bidi'] = 'bidi-override'
  },

  caps(style, obj) {
    style['text-transform'] = obj.val === 'false' ? 'none' : 'uppercase'
  },

  i(style, obj) {
    style['font-style'] = obj.val === 'false' ? 'normal' : 'italic'
  },

  shd(style, obj) {
    // TODO: w:val & w:color implementation
    style['background-color'] = `#${obj.fill}`
  },

  smallCaps(style, obj) {
    style['font-variant'] = obj.val === 'false' ? 'normal' : 'small-caps'
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
    return Object.entries(this.style).map(e => `${e[0]}: ${e[1]}${this._important ? ' !important' : ''};`).join(' ')
  }
}

module.exports = RunFormat