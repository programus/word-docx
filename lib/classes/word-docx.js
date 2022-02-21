const JSZipUtils = require('jszip-utils')
const Document = require('./document')

class WordDocx extends HTMLElement {
  constructor() {
    super()
    this._varName = undefined
    this.wordDoc = new Document(this)
    this.sharedAttributes = [
      'enable-css',
    ]
  }

  connectedCallback() {
    this._shadowRoot = this.noShadow ? this : this.attachShadow({mode: 'open'})
    this.root = document.createElement('div')
    this.root.style.position = 'relative'
    this.root.id = '_word-docx-root'
    this.sharedAttributes.forEach(attr => this.root.setAttribute(attr, this.getAttribute(attr)))
    this.style.position = 'relative'
    this._shadowRoot.append(this.root)
    this._generateStyle()
  }

  _generateStyle() {
    this.styleSheet = new CSSStyleSheet();
    (this.noShadow ? document : this._shadowRoot).adoptedStyleSheets = [this.styleSheet]
    this.addEventListener('css-class-updated', (evt) => {
      const styleSheetText = Object.entries(this.wordDoc.classes).map(e => `#_word-docx-root[enable-css='true'] .${e[0]} {${e[1].cssText}}`).join(';')
      this.styleSheet.replace(styleSheetText)
    })
  }

  static get observedAttributes() {
    return [
      'data-file',
      'var-name',
      'enable-css',
      'src',
    ]
  }

  attributeChangedCallback(name, _oldVal, newVal) {
    this[name.replace(/-\w/g, (v) => v.substring(1).toUpperCase())] = newVal
  }

  get noShadow() {
    return this.hasAttribute('no-shadow')
  }

  get varName() {
    return this._varName
  }

  set varName(value) {
    if (this._varName) {
      delete window[this._varName]
    }

    this._varName = value
    if (value) {
      window[this._varName] = this
    }
  }

  /**
   * @param {string} value
   */
  set enableCss(value) {
    this.root && this.root.setAttribute('enable-css', value)
  }

  /**
   * @param {string} value
   */
  set dataFile(value) {
    const fileCtrl = document.querySelector(value)
    fileCtrl.onchange = (evt) => {
      this.loadFile(evt.target.files[0])
    }
  }

  /**
   * @param {any} value
   */
  set src(value) {
    JSZipUtils.getBinaryContent(value, (err, data) => {
      if (err) {
        throw err
      } else {
        this.loadFile(data)
      }
    })
  }

  async loadFile(fileData) {
      await this.wordDoc.loadAsync(fileData)
      this.root.style.width = `${this.wordDoc.pageSize.w}pt`
      this.root.style.height = `${this.wordDoc.pageSize.h}pt`
      this.root.replaceChildren(this.wordDoc.bodyNode.htmlNode)
  }

  /**
   * Override selection format
   * @param {OverrideFormat} overrideFormat - the format to be used
   */
  overrideSelectionFormat(overrideFormat) {
    const selection = this._shadowRoot.getSelection ? this._shadowRoot.getSelection() : window.getSelection()
    if (selection && selection.toString().length > 0) {
      this.wordDoc.overrideFormat({
        selection, 
        format: overrideFormat,
      })
    }
  }
}

module.exports = WordDocx
