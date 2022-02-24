const RunFormat = require('../../run-format')

module.exports = (docNode, general) => {
  general(docNode, 'span')
  const props = docNode.doc.xpathSelect('w:rPr', docNode.xmlNode)[0]
  if (props) {
    const rf = new RunFormat(Array.from(props.children).filter(p => !p.$useClass))
    rf.styles.forEach(style => {
      const childNodes = docNode.htmlNode.childNodes
      const span = document.createElement('span')
      span.style.cssText = style.cssText
      Array.from(childNodes).forEach(node => span.appendChild(node))
      docNode.htmlNode.appendChild(span)
    })
    docNode.htmlNode.className = (props.$cssClasses || []).join(' ')
  }
}
