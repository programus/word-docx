const { generateHtmlNode } = require('../helpers/html-generator')

class DocNode {
  constructor(doc, xmlNode, parent) {
    this.doc = doc
    this.xmlNode = xmlNode
    this.xmlNode.$node = this
    this.tagName = this.xmlNode.tagName
    this.children = this.parseChildren()
    this.parent = parent
    this.dirty = true
    this.render()
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
    this.xmlNode.insertBefore(newNode.xmlNode, node.xmlNode)
    this.htmlNode.insertBefore(newNode.htmlNode, node.htmlNode)
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
    this.xmlNode.removeChild(node.xmlNode)
    this.htmlNode.removeChild(node.htmlNode)
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
