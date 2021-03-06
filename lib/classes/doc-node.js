const { generateHtmlNode } = require('../helpers/html-generator')

class DocNode {
  constructor(doc, xmlNode, parent) {
    this.doc = doc
    this.xmlNode = xmlNode
    this.xmlNode.$node = this
    this.tagName = this.xmlNode.tagName
    this.children = this.parseChildren()
    if (this.tagName == 'w:p') {
      this.attachInfo2T()
    }
    this.parent = parent
    this.dirty = true
    this.render()
  }

  attachInfo2T() {
    const tNodes = this.doc.xpathSelect('.//w:t', this.xmlNode)
    for (let i = 0; i < tNodes.length; i++) {
      const tXmlNode = tNodes[i]
      const tNode = tXmlNode.$node
      tNode.tIndex = i
      tNode.tStartPos = 0
      tNode.tEndPos = tXmlNode.textContent.length
    }
  }

  parseChildren() {
    return Array.from(this.xmlNode.children).map(n => new DocNode(this.doc, n, this))
  }

  render() {
    generateHtmlNode(this)
  }

  cloneNode(deep) {
    return new DocNode(this.doc, this.xmlNode.cloneNode(deep), undefined)
  }

  childIndex(nodeOrIndex) {
    let index = 0
    if (typeof(nodeOrIndex) === 'number') {
      index = parseInt(nodeOrIndex)
    } else {
      for (let i = 0; i < this.children.length; i++) {
        if (nodeOrIndex === this.children[i]) {
          index = i
          break
        }
      }
    }
    return index
  }

  remove() {
    if (this.parent) {
      this.parent.removeChild(this)
    }
    this.parent = undefined
  }

  insertBefore(newNode, referenceNodeOrIndex) {
    const index = this.childIndex(referenceNodeOrIndex)
    const node = this.children[index]
    newNode.remove()
    newNode.parent = this
    this.children.splice(this.childIndex(referenceNodeOrIndex), 0, newNode)
    this.xmlNode.insertBefore(newNode.xmlNode, node && node.xmlNode)
    this.htmlNode.insertBefore(newNode.htmlNode, node && node.htmlNode)
    return newNode
  }

  appendChild(node) {
    node.remove()
    node.parent = this
    this.children.push(node)
    this.xmlNode.appendChild(node.xmlNode)
    this.htmlNode.appendChild(node.htmlNode)
    return node
  }

  removeChild(nodeOrIndex) {
    const index = this.childIndex(nodeOrIndex)
    const node = this.children[index]
    this.xmlNode.removeChild(node && node.xmlNode)
    this.htmlNode.removeChild(node && node.htmlNode)
    this.children.splice(this.childIndex(nodeOrIndex), 1)
    node.parent = undefined
    return node
  }

  replaceChild(newChild, oldChild) {
    const index = this.childIndex(oldChild)
    this.removeChild(index)
    this.insertBefore(newChild, index)
    return oldChild
  }

  get firstChild() {
    return this.children[0]
  }

  get lastChild() {
    return this.children && this.children.length > 0 ? this.children[this.children.length - 1] : undefined
  }

  get previousSibling() {
    return this.xmlNode.previousSibling.$node
  }

  get nextSibling() {
    return this.xmlNode.nextSibling.$node
  }

  get indexInSameTags() {
    if (this.parent) {
      const brothers = this.parent.children.filter(n => n.tagName === this.tagName)
      for (let i = 0; i < brothers; i++) {
        if (brothers[i] === this) {
          return i
        }
      }
    }
  }

  get index() {
    if (this.parent) {
      const brothers = this.parent.children
      for (let i = 0; i < brothers; i++) {
        if (brothers[i] === this) {
          return i
        }
      }
    }
  }

  get html() {
    return this.htmlNode.outerHTML
  }

  get xml() {
    return this.xmlNode.toString()
  }
}

module.exports = DocNode
