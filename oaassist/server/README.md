## Welcome to OAAssist WebDemo

### 这个项目是什么？

这个工程为OA助手演示web端页面，旨在帮助大家能够快速理解并熟悉WPS加载项机制以及和浏览器调用交互的流程。
有些流程需要服务端的支持，因此用nodejs接口写了一个本地的模拟服务器。

### 工程结构

* StartupServer.js          nodejs脚本的服务端程序
* demo.html 				web端页面入口
* wwwroot					web服务根目录


### 操作流程
1.确保安装了支持wps加载项的企业版安装包。
2.确保安装了nodejs。
3.打开命令行，cd到本目录，执行“npm install express urlencode formidable”, 用于安装相关依赖包。
4.执行命令“node StartupServer.js”, 用于启动本地服务端程序。
5.浏览器打开demo.html, 开始体验相关流程。

### 注意事项

* 本工程只是演示demo
* 我们建议您结合具体的应用场景修改示例代码，这样更能够体现OA助手集成的应用场景
* 为了保护代码，建议代码上线前进行混淆
* 使用该工程的时候，必须要安装WPS专业版，请咨询QQ：3253920855
