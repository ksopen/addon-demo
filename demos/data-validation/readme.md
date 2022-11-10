## 目的
对表格中的身份证进行校验
1. 对身份证的合法性进行校验, 不合法的标记为黄色
2. 提取出身份信息(身份,性别,生日)

## 启动
1. 下载app-jssdk.zip包，解压后导入项目
2. 使用http-serve
```js
// 如果没有npm指令，请安装node
npm install -g http-server
```
3. 运行
打开命令行工具，切换到要开服务的目录下，执行
```js
http-server -p 8082
```
4. 查看
访问http://localhost:8082

