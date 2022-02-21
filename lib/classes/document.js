const JSZip = require('jszip')
const { isBrowser } = require('browser-or-node')
const xpath = require('xpath')

const DocNode = require('./doc-node')
const RunFormat = require('../helpers/run-format')

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
      'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml': 'numbering',
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
    const tf = new RunFormat(formatJson, true)
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

  /**
   * The description of a node
   * @typedef {Object} NodeDescription
   * @property {string} paraId - the paraId of the XML Node
   * @property {string} countTag - the tag used to count
   * @property {number} tTagIndex - the index of the node
   */

  /**
   * The information to find a node
   * @typedef {Object} NodeInfo
   * @property {Element} [xml] - the XML Element
   * @property {Element} [html] - the HTML Element
   * @property {NodeDescription} [info] - the information of the node
   */

  /**
   * The information of an edge of a range
   * @typedef {Object} EdgePosition
   * @property {NodeInfo} node - the information of the node
   * @property {number} textOffset - the text offset of the position
   */

  /**
   * The format style. 
   * The style included in the style is based on [the WordProcessingML definition](http://officeopenxml.com/WPtextFormatting.php)
   * The key of every style is the tag name without namespace and the value is attributes as a dictionary.
   * Also the attribute name without namespace is the key and the attribute value is the value.
   * @typedef {Object.<string, Object.<string, string>} FormatStyle
   */

  /**
   * The format
   * @typedef {Object} OverrideFormat
   * @property {[string]} [cssClasses] - the classes to be added, the class could be registered by {@link setCssClass}
   * @property {FormatStyle} [styles] - the styles to be added
   */

  /**
   * The information for overriding format
   * @typedef {Object} OverrideFormatParam
   * @property {EdgePosition} [from] - the start position
   * @property {EdgePosition} [to] - the end position
   * @property {Selection} [selection] - instance of [Selection](https://developer.mozilla.org/zh-CN/docs/Web/API/Selection) returned by [getSelection](https://developer.mozilla.org/zh-CN/docs/Web/API/Document/getSelection)
   * @property {OverrideFormat} format - the format to be used
   * 
   */

  /**
   * Override the format using specified params
   * @param {OverrideFormatParam} params - the params
   */
  overrideFormat(params) {
    const { from, to, format, selection } = params
    const { nodesNeedFormat, nodesNeedRenderOnly } = selection ? this.pickupRunNodesFromSelection(selection) : this.seekRunNodes(from, to)
    nodesNeedFormat.forEach(n => this.applyFormat(n, format))
    nodesNeedRenderOnly.forEach(n => n.render())
  }

  /**
   * Apply specified format to specified node
   * @param {DocNode} node
   * @param {OverrideFormat} format
   */
  applyFormat(node, format) {
    node.dirty = true
    const xmlNode = node.xmlNode
    const prNode = this.xpathSelect('w:rPr', xmlNode)[0] || this.document.createElement('w:rPr')
    if (!prNode.$node) {
      const docNode = new DocNode(node.doc, prNode)
      node.insertBefore(docNode, 0)
    }
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

  /**
   * Return all run nodes (w:r) between fromPair[1] (included) and toPair[1] (excluded)
   * @param {EdgePosition} [from] - the start position
   * @param {EdgePosition} [to] - the end position
   * @returns {{
   *    nodesNeedFormat: [DocNode],
   *    nodesNeedRenderOnly: [DocNode],
   * }} all run nodes need format or render
   */
  seekRunNodes(from, to) {
    const fromPair = this.findRunNodePair(from)
    const toPair = this.findRunNodePair(to, fromPair[0] && from.node.info && to.node.info && from.node.info.paraId === to.node.info.paraId ? fromPair : undefined)
    const nodesNeedFormat = []
    if (fromPair[1]) {
      const allXmlRunNodes = this.xpathSelect('//w:r', this.bodyNode.xmlNode)
      let mark = false
      let switchCount = 0
      for (let i = 0; i < allXmlRunNodes.length; i++) {
        const xmlNode = allXmlRunNodes[i]
        if (xmlNode.$node === fromPair[1] || xmlNode.$node === toPair[1]) {
          mark = !mark
          switchCount++
        }
        if (mark) {
          nodesNeedFormat.push(xmlNode.$node)
        }
        if (switchCount > 1) {
          break
        }
      }
    }
    return {
      nodesNeedFormat,
      nodesNeedRenderOnly: fromPair.concat(toPair).filter(n => n),
    }
  }

  /**
   * Return all run nodes (w:r) in the specified selection. Will split nodes while necessary.
   * @param {Selection} selection - instance of [Selection](https://developer.mozilla.org/zh-CN/docs/Web/API/Selection) returned by [getSelection](https://developer.mozilla.org/zh-CN/docs/Web/API/Document/getSelection)
   * @returns {{
   *    nodesNeedFormat: [DocNode],
   *    nodesNeedRenderOnly: [DocNode],
   * }} all run nodes need format or render
   */
  pickupRunNodesFromSelection(selection) {
    const nodesNeedFormat = []
    const nodesNeedRenderOnly = []
    if (selection.toString().length > 0) {
      const allXmlRunNodes = this.xpathSelect('//w:r', this.bodyNode.xmlNode)
      const possibleNodes = allXmlRunNodes.filter(node => selection.containsNode(node.$node.htmlNode, true)).map(n => n.$node)
      let fromOffset, toOffset, fromNode, toNode
      const positionComparison = selection.anchorNode.compareDocumentPosition(selection.focusNode)
      if (positionComparison & Node.DOCUMENT_POSITION_FOLLOWING) {
        fromOffset = selection.anchorOffset
        fromNode = selection.anchorNode.parentElement.$node
        toOffset = selection.focusOffset
        toNode= selection.focusNode.parentElement.$node
      } else if (positionComparison & Node.DOCUMENT_POSITION_PRECEDING) {
        fromOffset = selection.focusOffset
        fromNode = selection.focusNode.parentElement.$node
        toOffset = selection.anchorOffset
        toNode = selection.anchorNode.parentElement.$node
      } else {
        fromOffset = Math.min(selection.anchorOffset, selection.focusOffset)
        fromNode = selection.focusNode.parentElement.$node
        toOffset = Math.max(selection.anchorOffset, selection.focusOffset)
        toNode = selection.anchorNode.parentElement.$node
      }
      if (possibleNodes.length > 0) {
        const fromPair = this.splitNode(fromNode, fromOffset)
        const toPair = this.splitNode(toNode, toOffset, possibleNodes.length == 1 ? fromPair : undefined)
        nodesNeedFormat.push(...[fromPair[1]].concat(possibleNodes.slice(1, possibleNodes.length - 1)).concat(toPair[0]).filter(n => n))
        nodesNeedRenderOnly.push(...fromPair.concat(toPair).filter(n => n))
      }
    }

    return {
      nodesNeedFormat,
      nodesNeedRenderOnly,
    }
  }

  /**
   * Return the two nodes splitted by the node information
   * @param {EdgePosition} nodeInfo - the position
   * @param {[DocNode|undefined, DocNode|undefined]} [referPair] - the possible start node pair for reference
   * @returns {[DocNode|undefined, DocNode|undefined]} the node pair: [0] - the privious node of the position, [1] - the next node of the position, both may be undefined
   */
  findRunNodePair(nodeInfo, referPair) {
    const node = this.findNode(nodeInfo)
    return this.splitNode(node, nodeInfo.textOffset, referPair)
  }

  /**
   * 
   * @param {DocNode|undefined} node - the node to be splitted
   * @param {number} textOffset - the text offset of the position to be splitted
   * @param {[DocNode|undefined, DocNode|undefined]} [referPair] - the possible start node pair for reference
   * @returns {[DocNode|undefined, DocNode|undefined]} splitted nodes, there are garantee two elements in the array
   */
  splitNode(node, textOffset, referPair) {
    const originalText = node && node.xmlNode && node.xmlNode.textContent
    const runNode = WordDocument.findRunNode(node)
    const rangeInSameRunNode = referPair && referPair[1] === runNode
    if (rangeInSameRunNode) {
      textOffset -= referPair.textOffset || 0
    }
    const ret = [undefined, runNode]
    ret.textOffset = textOffset
    if (textOffset > 0 && originalText && originalText.length > 0) {
      node.xmlNode.textContent = originalText.substring(0, textOffset)
      node.dirty = true
      runNode.dirty = true
      const newNode = runNode.cloneNode(true)
      newNode.tEndPos -= newNode.tEndPos - textOffset
      node.xmlNode.textContent = originalText.substring(textOffset)
      node.tStartPos += textOffset - node.tStartPos
      runNode.parent.insertBefore(newNode, runNode)
      ret[0] = newNode
      if (rangeInSameRunNode) {
        referPair[1] = newNode
      }
    }
    return ret
  }

  /**
   * Return the closest run tag of the current node
   * @param {DocNode|undefined} node - the node
   * @returns {DocNode|undefined} - the run node or undefined
   */
  static findRunNode(node) {
    return !node || node.tagName === 'w:r' ? node : WordDocument.findRunNode(node.parent)
  }

  /**
   * 
   * Find the node by using the provided node information
   * @param {EdgePosition} nodeInfo - an Object in format
   * @returns {DocNode|undefined} the node found
   */
  findNode(nodeInfo) {
    const textOffset = nodeInfo.textOffset
    const { xml, html, info } = nodeInfo.node
    let node = (xml || html || {}).$node
    if (!node && info) {
      const { paraId, countTag, tTagIndex, selector, xpath } = info
      let mlNode
      if (selector) {
        mlNode = this.bodyNode.htmlNode.querySelector(selector)
      }
      if (!mlNode && xpath) {
        mlNode = this.xpathSelect(xpath, this.bodyNode.xmlNode)[0]
      }
      if (!mlNode) {
        mlNode = this.xpathSelect(`//w:p[@w14:paraId="${paraId}"]//w:t`, this.bodyNode.xmlNode)
                     .filter(n => n.$node.tIndex === tTagIndex && n.$node.tStartPos <= textOffset && n.$node.tEndPos > textOffset)[0]
      }
      if (mlNode) {
        node = mlNode.$node
      }
    }
    return node
  }
}

module.exports = WordDocument
