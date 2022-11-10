const path = require('path')   //调用路径
const HtmlWebpackPlugin = require('html-webpack-plugin');
const FileManagerPlugin = require('filemanager-webpack-plugin');

module.exports = {
  mode: 'development',    //开发模式
  entry: "./index.js",
  target: 'web',
  output: { // 出口路径
    path: path.resolve(__dirname, 'dist'),
    // [hash]让每一次生成的文件都带上HASH值
    filename: 'index.js'
  },
  devServer: {
    port: 8001, // 端口号
    open: true, // 编译完成会自动打开浏览器
    proxy: {
      '/api': {
        target: 'https://www.maomin.club',
        pathRewrite: { '^/api': '' },
        secure: false,
        changeOrigin: true
      }
    }
  },
  // 使用插件
  plugins: [
    new HtmlWebpackPlugin({
      // 指定要编译的文件，不指定的话会按照默认的模板创建一个html
      template: './index.html',
      // 编译完成输出的文件名
      filename: 'index.html',
      // 给引入的文件通过问号传参设置HASH戳（清除缓存的），但是真实项目我们一般都是每一次编译生成不同的JS文件引入，详见上面的出口路径设置
      // hash: true,
      // 控制压缩
      minify: {
        collapseWhitespace: true, // 干掉空格
        removeComments: true, // 干掉注释
        removeAttributeQuotes: true, // 干掉双引号
        removeEmptyAttributes: true // 干掉空属性
      }
    }),
    new FileManagerPlugin({  //初始化 filemanager-webpack-plugin 插件实例
      events: {
        onEnd: {
          delete: [   //首先需要删除项目根目录下的dist.zip
            './dist.zip',
          ],
          archive: [
            {source: './dist', destination: './dist.zip'},
          ]
        }
      }
    })
  ],
  // 使用loader加载器来处理规则
  module: {
    rules: [{
      // 基于正则匹配处理哪些文件
      test: /\.(css)$/i,
      // 控制使用哪个加载器loader（有顺序的：数组从右到左执行）
      use: [
        "style-loader", // 把编译好的css插入到页面的HEAD中（内嵌式样式）
        "css-loader" // 编译@import/url()这种语法的
      ]
    }]
  }
}
