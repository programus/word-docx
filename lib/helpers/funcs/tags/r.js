const TextFormat = require('../../text-format')

module.exports = (docNode, general) => {
  general(docNode, 'span')
  const props = docNode.doc.xpathSelect('w:rPr', docNode.xmlNode)[0]
  if (props) {
    const tf = new TextFormat(Array.from(props.children).filter(p => !p.$useClass))
    docNode.htmlNode.style.cssText = tf.cssText
    docNode.htmlNode.classList = props.$cssClasses || []
  }
}
