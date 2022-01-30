// generate content of this file
if (require.main === module) {
  const fs = require('fs')
  const path = require('path')
  const funcsBody = fs.readdirSync(path.join(__dirname, 'tags')).map(f => {
    const tagName = f.replace(/\.js$/g, '')
    return `  'w:${tagName}': require('./tags/${tagName}'),`
  }).join('\n')

  const content = fs.readFileSync(__filename).toString().replace(
    /(\/\/ g\{)(?:.|[\n\r])*(\/\/ \})/gm, 
    `$1\n${funcsBody}\n$2`)
  fs.writeFileSync(__filename, content)
  process.exit(0)
}

const funcs = {
// g{
  'w:br': require('./tags/br'),
  'w:cr': require('./tags/cr'),
  'w:noBreakHyphen': require('./tags/noBreakHyphen'),
  'w:p': require('./tags/p'),
  'w:r': require('./tags/r'),
  'w:t': require('./tags/t'),
  'w:tab': require('./tags/tab'),
// }
}

module.exports = funcs
