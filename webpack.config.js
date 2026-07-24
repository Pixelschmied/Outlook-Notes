/* eslint-disable */
const path = require('path')
const webpack = require('webpack')
const HtmlWebpackPlugin = require('html-webpack-plugin')
const CopyWebpackPlugin = require('copy-webpack-plugin')
const pkg = require('./package.json')

/**
 * Two HTML entry points, as an Outlook task-pane add-in expects:
 *  - taskpane: the notepad UI that docks next to the mail
 *  - commands: the function-file the ribbon button loads
 *
 * office.js itself is intentionally NOT bundled. Microsoft requires it to be
 * loaded from the official CDN at runtime (see the <script> tag in the HTML),
 * so we never redistribute it here.
 */
module.exports = (env, argv) => {
  const dev = argv.mode !== 'production'
  return {
    devtool: dev ? 'source-map' : false,
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: 'ts-loader',
          exclude: /node_modules/
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'assets', to: 'assets' },
          { from: 'src/taskpane/taskpane.css', to: 'taskpane.css' },
          { from: 'manifest.xml', to: 'manifest.xml' }
        ]
      })
    ],
    devServer: {
      static: { directory: path.join(__dirname, 'dist') },
      server: 'https',
      port: 3000,
      hot: true,
      headers: { 'Access-Control-Allow-Origin': '*' }
    }
  }
}
