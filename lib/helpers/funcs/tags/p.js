module.exports = (docNode, general) => {
  general(docNode, 'div')
  docNode.htmlNode.style.minHeight = '1em'
}