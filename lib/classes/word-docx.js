const Document = require('./document')

class WordDocx extends HTMLElement {
  constructor() {
    super()
    this._varName = undefined
    this.wordDoc = new Document(this)
  }

  connectedCallback() {
    this._shadowRoot = this.useShadow ? this.attachShadow({mode: 'open'}) : this
    this.root = document.createElement('div')
    this.root.style.position = 'relative'
    this.root.id = '_word-docx-root'
    this._shadowRoot.style.position = 'relative'
    this._shadowRoot.append(this.root)
    this._generateStyle()
  }

  _generateStyle() {
    this.styleSheet = new CSSStyleSheet();
    (this.useShadow ? this._shadowRoot : document).adoptedStyleSheets = [this.styleSheet]
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
    ]
  }

  attributeChangedCallback(name, _oldVal, newVal) {
    this[name.replace(/-\w/g, (v) => v.substring(1).toUpperCase())] = newVal
  }

  get useShadow() {
    return this.hasAttribute('use-shadow')
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
    this.root.setAttribute('enable-css', value)
  }

  /**
   * @param {string} value
   */
  set dataFile(value) {
    const fileCtrl = document.querySelector(value)
    fileCtrl.onchange = async (evt) => {
      await this.wordDoc.loadAsync(evt.target.files[0])
      this.root.style.width = `${this.wordDoc.pageSize.w}px`
      this.root.style.height = `${this.wordDoc.pageSize.h}px`
      this.root.replaceChildren(this.wordDoc.bodyNode.htmlNode)
    }
  }
}

module.exports = WordDocx
