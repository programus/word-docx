const path = require('path')
const WebpackShellPluginNext = require('webpack-shell-plugin-next')

module.exports = {
  entry: './lib/index.js',
  output: {
    filename: process.env.NODE_ENV === 'development' ? 'word-docx.dev.js' : 'word-docx.min.js',
    path: path.resolve(__dirname, 'dist'),
  }, 
  externals: {
    jsdom: 'jsdom',
    fs: 'fs',
    path: 'path',
  },
  plugins: [
    new WebpackShellPluginNext({
      onBuildStart: {
        scripts: [
          'node lib/helpers/funcs/index.js',
        ],
        blocking: true,
        parallel: false,
      },
    }),
  ],
}