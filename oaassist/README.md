## Welcome to WPS OAAssist Demo

### 这个项目是什么？

    这个工程主要提供常见的OA助手的场景示例来演示网页端启动WPS客户端并和WPS加载项交互WPS API的功能，方便大家能够快速理解并熟悉WPS加载项机制以及和浏览器调用交互的流程。

### 工程结构

* demo.html 	包含了本地是否安装了正确的wps安装包、是否启动了本地服务端等的环境检测。
* server 	包含了一些前端文件和演示场景的模板文件，为网页端场景代码, 此外有几个场景需要服务端的支持，用nodejs写了一个本地服务程序用于模拟服务端场景。
* EtOAAssist	WPS 表格组件的OA助手WPS加载项，提供简单的OA场景功能示例。（单独的网页项目）
* WppOAAssist	WPS 演示组件的OA助手WPS加载项，提供简单的OA场景功能示例。（单独的网页项目）
* WpsOAAssist	WPS 文字组件的OA助手WPS加载项，提供常见的OA场景功能示例。（单独的网页项目）
### demo启动
1. 安装WPS,WPS版本支持情况

    WPS Win：企业版：11.8.2.8808；个人版：11.1.0.9566

    Linux 企业版：11.8.2.9346 ； 个人版暂不支持

    他们之后的版本，含他们自己

    这些版本是稳定支持的，之前的2019版本也支持，不推荐用了，jsapi支持的不稳定。

2. 安装node(仅demo需要)
    [windows安装](https://www.cnblogs.com/liuqiyun/p/8133904.html)

    [Linux安装](https://www.cnblogs.com/sirdong/p/11447739.html)
    
    使用node的作用是：

    * 静态资源转发。/plugin/et指向EtOAAssist目录，/plugin/wps指向WpsOAAssist目录，/plugin/et指向WppOAAssist目录。以及file目录下文档的访问。

    * 后端接口提供：提供了下载文件接口/Download/文件名 和 上传文件接口/Upload

    在实际项目中，不需要安装node，静态资源转发由tomcat、nginx或者其他中间件实现。后端接口由java或者php语言实现。

3. 进入server目录下
4. npm config set registry http://registry.npm.taobao.org //切换npm淘宝镜像源
5. npm install //安装相应依赖
6. node StartupServer //启动demo的服务


### WPS重要地址

* WPS配置文件oem.ini地址
```
    oem.ini目录地址：
    windows:
        1. 安装路径\WPS Offlce\一串数字（版本号）\offlce6\cfgs\
        2. 鼠标右键点击左面的wps文字图标==>打开文件位置==>在同级目录中找到cfgs目录
    linux:
        普通linux操作系统：
             /opt/kingsoft/wps-office/office6/cfgs/
        uos操作系统:
            /opt/apps/cn.wps.wps-office-pro/files/kingsoft/wps-office/office6/cfgs/
```


* 加载项管理文件存放位置（jsaddons目录）
```
    jsaddons目录地址：
    windows:
        我的电脑地址栏中输入：%appdata%\kingsoft\wps\jsaddons
    linux:
        我的电脑地址栏中输入：~/.local/share/Kingsoft/wps/jsaddons

```

### 调试器开启和使用

    1. 配置oem.ini,在support栏下配置JsApiShowWebDebugger=true
    2. linux机器上需要使用quickstartoffice restart重启WPS
        普通linux操作系统：
            电脑终端执行quickstartoffice restart
        uos操作系统:
            电脑终端执行 cd /opt/apps/cn.wps.wps-office-pro/files/bin
            ./quickstartoffice restart
    3. WPS打开后，在有文档的情况下按alt+F12(index.html页面的调试器)
    4. ShowDialog和Taskpane页面的调试器，点击该弹窗或者任务窗格，按F12
    如果无法打开调试器，那么说明加载项加载失败了，排查加载项管理文件是否生成，加载项管理文件中的加载项地址是否正确



### 项目集成
1. 部署加载项

    * 将WpsOAAssist,EtOAAssist,WppOAAssist这三个目录分别部署到服务器上

    [部署到tomcat](https://jingyan.baidu.com/article/22a299b5c6cfb09e18376a62.html)
    [部署到nginx](https://www.cnblogs.com/amazingjava/p/13411644.html)
2. 配置加载项管理文件

     加载项有两种部署模式，publish模式和jsplugins.xml模式，**这两种模式是WPS去找到加载项管理文件的方式**，每个模式都有对应的管理文件,WPS启动时，会去jsaddons目录读取publish.xml和jsplugins.xml文件。
     
     
    * 区别：

        * 管理文件生成方式不一样

        publish模式是通过在网页中调用本地服务的端口，在客户本地jsaddons目录中生成publish.xml文件，https://kdocs.cn/l/cpOfxONhn8Yg [金山文档] publish自动安装加载项.docx

        jsplugins.xml模式是在oem.ini中配置好地址，在WPS启动时，会自动去服务端拉取地址指向的jsplugins.xml文件，放到客户本地的jsaddons目录中。在实际项目中，将jsplugins.xml文件地址告知我们，由我们将jsplugins打包进WPS安装包中，用户安装二次打包后的安装包即可使用
        
    * 相同
        * 都有离线和在线模式
        离线模式和在线模式是去根据加载项管理文件中的加载项地址，去拉取代码的方式，模式介绍请看文档
        
        https://kdocs.cn/l/cBk8tsBIf
        [金山文档] 加载项在线模式和离线模式.docx
        

### 注意事项

* 本工程只是演示demo
* 我们建议您修改示例代码结合具体的应用场景部署到服务器上面，这样更能够体现OA助手集成的应用场景
* 为了保护代码，建议代码上线前进行混淆
* 使用该工程的时候，必须要安装WPS专业版，请咨询QQ：3253920855


        
