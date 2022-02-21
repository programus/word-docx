const ParaFormat = require('../../para-format')

module.exports = (docNode, general) => {
  general(docNode, 'div')
  const pPrNode = docNode.doc.xpathSelect('w:pPr', docNode.xmlNode)[0]
  if (pPrNode) {
    const pf = new ParaFormat(pPrNode)
    docNode.htmlNode.style.cssText = pf.cssText
    docNode.htmlNode.classList = pf.classList
  }
  docNode.htmlNode.style.minHeight = '1em'
}