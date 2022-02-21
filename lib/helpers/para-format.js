const { object2CssText, addCssUnits } = require('./utils')

const processors = {
  /**
   * http://officeopenxml.com/WPalignment.php
   * @param {ParaFormat} pf 
   * @param {Element} e 
   */
  jc(pf, e) {
    const val = e.getAttribute('w:val')
    pf.style['text-align'] = val.replace(/both|distribute/g, 'justify')
    if (val === 'distribute') {
      pf.style['text-align-last'] = pf.style['text-align']
      pf.style['text-justify'] = 'inter-character'
    }
  },

  /**
   * http://officeopenxml.com/WPindentation.php
   * @param {ParaFormat} pf 
   * @param {Element} e 
   */
  ind(pf, e) {
    const left = e.getAttribute('w:left') || e.getAttribute('w:start')
    const right = e.getAttribute('w:right') || e.getAttribute('w:end')
    const indent = e.getAttribute('w:firstLine') || -e.getAttribute('w:hanging')
    Object.entries({left, right}).forEach(([k, v]) => {
      const key = `padding-${k}`
      pf.style[key] = addCssUnits(pf.style[key], `${v / 20}pt`)
    })
    if (indent) {
      pf.style['text-indent'] = `${indent / 20}pt`
    }
  },

  /**
   * http://officeopenxml.com/WPborders.php
   * @param {ParaFormat} pf 
   * @param {Element} e 
   */
  pBdr(pf, e) {
    const styleMap = {
      single: 'solid',
      dotted: 'dotted',
      dashed: 'dashed',
      double: 'double',
    }
    Array.from(e.children).forEach(child => {
      const name = child.tagName.replace(/\w+:(\w+)$/g, '$1')
      const style = styleMap[child.getAttribute('w:val')] || 'solid'
      const width = child.getAttribute('w:sz') || 0
      const color = child.getAttribute('w:color')
      const space = child.getAttribute('w:space')

      pf.style[`border-${name}-style`] = style
      if (width) {
        pf.style[`border-${name}-width`] = `${width / 8}pt`
      }
      if (color) {
        pf.style[`border-${name}-color`] = color
      }
      if (space) {
        const key = `padding-${name}`
        pf.style[key] = addCssUnits(pf.style[key], `${space}pt`)
      }
    })
  },
}

class ParaFormat {
  /**
   * 
   * @param {Element} pPrNode 
   */
  constructor(pPrNode) {
    this.style = {}
    this.classes = {}
    Array.from(pPrNode.children).forEach(e => (processors[e.tagName.substring('w:'.length)] || ((x) => {}))(this, e))
  }

  get classList() {
    return Object.entries(this.classes).filter(e => e[1]).map(e => e[0])
  }

  get cssText() {
    return object2CssText(this.style, this._important)
  }
}

module.exports = ParaFormat
