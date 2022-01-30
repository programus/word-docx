const JSZip = require('jszip')
const { isBrowser } = require('browser-or-node')
const xpath = require('xpath')

const DocNode = require('./doc-node')
const TextFormat = require('../helpers/text-format')

if (!isBrowser) {
  const { JSDOM } = require('jsdom')
  global.window = new JSDOM().window
  global.document = window.document
  global.DOMParser = window.DOMParser
}

class WordDocument {
  constructor(owner) {
    this.owner = owner
    this.domParser = new DOMParser()
    this.contentTypeMap = {
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml': 'document',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml': 'styles',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml': 'settings',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml': 'webSettings',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml': 'fontTable',
      'application/vnd.openxmlformats-officedocument.theme+xml': 'theme',
      'application/vnd.openxmlformats-package.core-properties+xml': 'coreProp',
      'application/vnd.openxmlformats-officedocument.extended-properties+xml': 'appProp',
    }
    this.classes = {}
  }

  async loadAsync(data) {
    this.zip = await JSZip.loadAsync(data)
    this.meta = this.domParser.parseFromString(await this.zip.file('[Content_Types].xml').async('string'), 'text/xml')

    const overrides = this.meta.getElementsByTagName('Override')

    await Promise.all(Array.from(overrides).map(node => {
      const contentType = node.getAttribute('ContentType')
      const propName = this.contentTypeMap[contentType]
      return {node, contentType, propName}
    }).filter(e => e.propName).map(async e => {
      const { node, contentType, propName } = e
      const filePath = node.getAttribute('PartName').substring(1)
      const s = await this.zip.file(filePath).async('string')
      this[propName] = this.domParser.parseFromString(s, 'text/xml')
    }))

    const docNode = this.document.children[0]
    this.docNamespaces = Object.values(docNode.attributes).filter(n => n.name.startsWith('xmlns:')).reduce((a, c) => {
      a[c.name.substring('xmlns:'.length)] = c.value
      return a
    }, {})
    this.xpathSelect = xpath.useNamespaces(this.docNamespaces)

    this.parseSection()
    this.parseBody()

    if (this.owner) {
      const event = new CustomEvent('file-loaded', {
        bubbles: true,
        composed: true,
        detail: {
          wordDoc: this,
        },
      })
      this.owner.dispatchEvent(event)
    }
  }

  get loaded() {
    return !!this.bodyNode
  }

  setCssClass(className, formatJson) {
    const tf = new TextFormat(formatJson, true)
    this.classes[className] = tf
    if (this.owner) {
      this.owner.dispatchEvent(new CustomEvent('css-class-updated', {
        bubbles: true,
        composed: true,
        detail: {
          className, 
          cssText: tf.cssText,
        },
      }))
    }
  }

  parseSection() {
    const sectPr = this.xpathSelect('//w:sectPr', this.document)[0]
    const pgSz = this.xpathSelect('w:pgSz', sectPr)[0]
    this.pageSize = {
      w: pgSz.getAttribute('w:w') / 20,
      h: pgSz.getAttribute('w:h') / 20,
    }
  }

  parseBody() {
    const body = this.xpathSelect('//w:body', this.document)[0]
    this.bodyNode = new DocNode(this, body)
  }

  overrideFormat(params) {
    const { from, to, format } = params
    const nodesNeedChange = this.seekRunNodes(this.findRunNodePair(from), this.findRunNodePair(to))
    nodesNeedChange.forEach(n => this.applyFormat(n, format))
  }

  applyFormat(node, format) {
    node.dirty = true
    const xmlNode = node.xmlNode
    const prNode = this.xpathSelect('w:rPr', xmlNode)[0] || xmlNode.insertBefore(this.document.createElement('w:rPr'), xmlNode.firstChild)
    prNode.$cssClasses = format.cssClasses
    let styles = (format.cssClasses || []).map(c => this.classes[c])
    if (format.styles) {
      styles.push(format.styles)
    }
    styles.forEach(formatJson => {
      Object.entries(formatJson).forEach(entry => {
        const tagName = entry[0]
        const tagNameNS = `w:${tagName}`
        const styleNode = this.xpathSelect(`.//${tagNameNS}`, prNode)[0] || prNode.appendChild(this.document.createElement(tagNameNS))
        if (!styleNode.$original) {
          styleNode.$original = {}
        }
        if (!styleNode.$original[tagName]) {
          // save original format
          styleNode.$original[tagName] = entry[1]
        }
        styleNode.$useClass = format.cssClasses && format.cssClasses.length > 0
        Object.entries(entry[1]).forEach(e => {
          styleNode.setAttribute(`w:${e[0]}`, e[1])
        })
      })
    })
    node.render()
  }

  seekRunNodes(fromPair, toPair) {
    const allXmlRunNodes = this.xpathSelect('//w:r', this.bodyNode.xmlNode)
    const ret = []
    let mark = false
    let switchCount = 0
    for (let i = 0; i < allXmlRunNodes.length; i++) {
      const xmlNode = allXmlRunNodes[i]
      if (xmlNode.$node === fromPair[1] || xmlNode.$node === toPair[1]) {
        mark = !mark
        switchCount++
      }
      if (mark) {
        ret.push(xmlNode.$node)
      }
      if (switchCount > 1) {
        break
      }
    }
    return ret
  }

  findRunNodePair(nodeInfo) {
    const node = this.findNode(nodeInfo.node)
    return this.splitNode(node, nodeInfo.textOffset)
  }

  splitNode(node, textOffset) {
    const originalText = node && node.xmlNode && node.xmlNode.textContent
    const runNode = WordDocument.findRunNode(node)
    const ret = [undefined, runNode]
    if (textOffset > 0 && originalText.length > 0) {
      node.dirty = true
      node.xmlNode.textContent = originalText.substring(0, textOffset)
      const newNode = runNode.cloneNode(true)
      node.xmlNode.textContent = originalText.substring(textOffset)
      runNode.parent.insertBefore(newNode, runNode)
      ret[0] = newNode
    }
    return ret
  }

  static findRunNode(node) {
    return node.tagName === 'w:r' || !node ? node : WordDocument.findRunNode(node.parent)
  }

  findNode(nodeInfo) {
    const { xml, html, info } = nodeInfo
    let node = (xml || html || {}).$node
    if (!node && info) {
      const { paraId, countTag, countIndex, selector, xpath } = info
      let mlNode
      if (selector) {
        mlNode = this.bodyNode.htmlNode.querySelector(selector)
      }
      if (!mlNode && xpath) {
        mlNode = this.xpathSelect(xpath, this.bodyNode.xmlNode)[0]
      }
      if (!mlNode) {
        mlNode = this.xpathSelect(`//w:p[@w14:paraId="${paraId}"]//w:${countTag || 't'}`, this.bodyNode.xmlNode)[countIndex]
      }
      if (mlNode) {
        node = mlNode.$node
      }
    }
    return node
  }
}

module.exports = WordDocument
