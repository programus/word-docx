module.exports = (docNode, general) =>  {
  general(docNode, 'span')
  docNode.htmlNode.innerHTML = '&emsp;'
}
