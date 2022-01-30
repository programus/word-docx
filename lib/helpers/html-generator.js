const funcs = require('./funcs')

function generateHtmlNode(docNode) {
  if (docNode.dirty) {
    f = funcs[docNode.tagName]
    f ? f(docNode, general) : general(docNode)
  }
  docNode.dirty = false

  docNode.children.forEach(c => generateHtmlNode(c))
}

function general(docNode, tagName = 'span') {
  if (!docNode.htmlNode) {
    const tag = document.createElement(tagName)
    tag.$node = docNode
    tag.setAttribute('xtag', docNode.tagName)
    Object.values(docNode.xmlNode.attributes).forEach(attr => tag.setAttribute(attr.name, attr.value))
    docNode.htmlNode = tag
    linkChildren(docNode)
  }
}

function linkChildren(docNode) {
  for (let c of docNode.children) {
    docNode.htmlNode.appendChild(c.htmlNode)
  }
}

module.exports = {
  generateHtmlNode,
}