var pluginsMode = location.search.split("=")[1];//截取url中的参数值
var pluginType = WpsInvoke.ClientType.wps//加载项类型wps,et,wpp
var pluginName = "WpsOAAssist";//加载项名称
var wpsClient = new WpsClient(pluginType);//初始化一个多进程对象，多进程时才需要
var clientStr = pluginName + pluginType + "ClientId"
//单进程封装开始
/**
 * 此方法是根据wps_sdk.js做的调用方法封装
 * 可参照此定义
 * @param {*} funcs         这是在WPS加载项内部定义的方法，采用JSON格式（先方法名，再参数）
 * @param {*} front         控制着通过页面执行WPS加载项方法，WPS的界面是否在执行时在前台显示
 * @param {*} jsPluginsXml  指定一个新的WPS加载项配置文件的地址,动态传递jsplugins.xml模式，例如：http://127.0.0.1:3888/jsplugins.xml
 * @param {*} isSilent      隐藏打开WPS，如果需要隐藏，那么需要传递front参数为false
 */


function _WpsInvoke(funcs, front, jsPluginsXml,isSilent) {
    var info = {};
    info.funcs = funcs;
    if(isSilent){//隐藏启动时，front必须为false
        front=false;
    }
    /**
     * 下面函数为调起WPS，并且执行加载项WpsOAAssist中的函数dispatcher,该函数的参数为业务系统传递过去的info
     */
    if (pluginsMode != 2) {//单进程
        singleInvoke(info,front,jsPluginsXml,isSilent)
    } else {//多进程
        multInvoke(info,front,jsPluginsXml,isSilent)
    }

}

//单进程
function singleInvoke(info,front,jsPluginsXml,isSilent){
    WpsInvoke.InvokeAsHttp(pluginType, // 组件类型
        pluginName, // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据        
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (result.status) {
                if (result.status == 100) {
                    WpsInvoke.AuthHttpesCert('请在稍后打开的网页中，点击"高级" => "继续前往"，完成授权。')
                    return;
                }
                alert(result.message)

            } else {
                console.log(result.response)
                showresult(result.response)
            }
        },
        front,
        jsPluginsXml,
        isSilent)

    /**
     * 接受WPS加载项发送的消息
     * 接收消息：WpsInvoke.RegWebNotify（type，name,callback）
     * WPS客户端返回消息： wps.OAAssist.WebNotify（message）
     * @param {*} type 加载项对应的插件类型
     * @param {*} name 加载项对应的名字
     * @param {func} callback 接收到WPS客户端的消息后的回调函数，参数为接受到的数据
     */
    WpsInvoke.RegWebNotify(pluginType, pluginName, handleOaMessage)
}
//多进程
function multInvoke(info,front,jsPluginsXml,isSilent){
    wpsClient.jsPluginsXml = jsPluginsXml ? jsPluginsXml : "https://127.0.0.1:3888/jsplugins.xml";
    if (localStorage.getItem(clientStr)) {
        wpsClient.clientId = localStorage.getItem(clientStr)
    }
    if(isSilent){
        wpsClient.StartWpsInSilentMode(pluginName,function(){//隐藏启动后的回调函数
            mult(info,front)
        })
    }else{
        mult(info,front)
    }
    wpsClient.onMessage = handleOaMessage
}
//多进程二次封装
function mult(info,front){
    wpsClient.InvokeAsHttp(
        pluginName, // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据        
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (wpsClient.clientId) {
                localStorage.setItem(clientStr, wpsClient.clientId)
            }
            if (result.status !== 0) {
                console.log(result)
                if (result.message == '{\"data\": \"Failed to send message to WPS.\"}') {
                    wpsClient.IsClientRunning(function (status) {
                        console.log(status)
                        if (status.response == "Client is running.")
                            alert("任务发送失败，WPS 正在执行其他任务，请前往WPS完成当前任务")
                        else {
                            wpsClient.clientId = "";
                            wpsClient.notifyRegsitered = false;
                            localStorage.setItem(clientStr, "")
                            mult(info)
                        }
                    })
                    return;
                }
                else if (result.status == 100) {
                    // WpsInvoke.AuthHttpesCert('请在稍后打开的网页中，点击"高级" => "继续前往"，完成授权。')
                    return;
                }
                alert(result.message)
            } else {
                console.log(result.response)
            }
        },
        front)
}
function handleOaMessage(data) {
    console.log(data)
}

function handleOaFunc1(message) {
    alert("我是函数handleOaFunc1，我接收到的参数是：" + message)
}
function handleOaFunc2(message) {
    alert("我是函数handleOaFunc2，我接收到的参数是：" + message)
    var span = window.parent.document.getElementById("webnotifyspan")
    span.innerHTML = message
}
/**
 * 处理WPS加载项的方法返回值
 *
 * @param {*} resultData
 */
function showresult(resultData) {
    let json = eval('(' + resultData + ')')
    switch (json.message) {
        case "GetDocStatus": {
            let docstatus = json.docstatus
            if (typeof docstatus != "undefined") {
                let str = "文档保存状态：" +
                    docstatus.saved +
                    "\n文档字数：" +
                    docstatus.words +
                    "\n文档页数：" + docstatus.pages
                alert(str)
            }
        }
    }
}
/**
 * 这是页面中针对代码显示的变量定义，开发者无需关心
 */
var _wps = {}

// 此处往下，都是对于前端页面如何调用WPS加载项方法的样例，开发者请参考

function newDoc() {
    _WpsInvoke([{
        "NewDoc": {}
    }],
        true,
        "http://127.0.0.1:8080/iestart/jsplugins.xml") // NewDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['newDoc'] = {
    action: newDoc,
    code: _WpsInvoke.toString() + "\n\n" + newDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，打开WPS文字后,新建一个空白doc文档\n\
\n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行插件中的js函数NewDoc,新建一个空白doc\n\
    funcs参数说明:\n\
        NewDoc方法对应于OA助手dispatcher支持的方法名\n\
"
}

function GetDemoPath(fileName) {

    var url = document.location.host;
    return document.location.protocol + "//" + url + "/file/" + fileName;
}

function GetUploadPath() {
    var url = document.location.host;
    return document.location.protocol + "//" + url + "/Upload";
}

function GetDemoPngPath() {
    var url = document.location.toString();
    url = decodeURI(url);
    if (url.indexOf("/") != -1) {
        url = url.substring(0, url.lastIndexOf("/"));
    }
    if (url.length !== 0)
        url = url.concat("/WPS.png");

    if (!String.prototype.startsWith) {
        String.prototype.startsWith = function (searchString, position) {
            position = position || 0;
            return this.indexOf(searchString, position) === position;
        };
    }

    if (url.startsWith("file:///"))
        url = url.substr("file:///".length);
    return url;
}

function openDoc() {

    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())
    var uploadFieldName = prompt("请输入文档上传到业务系统时自定义字段：", "自定义字段")
    var backupPath = prompt("请输入文档备份路径:")

    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "uploadFieldName": uploadFieldName,
            "picPath": GetDemoPngPath(),
            "copyUrl": backupPath,
            "userName": "东方不败"
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['openDoc'] = {
    action: openDoc,
    code: _WpsInvoke.toString() + "\n\n" + openDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径下载并打开对应的文档，保存将自动上传指定服务器地址\n\
\n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            uploadFieldName 文档上传到业务系统时自定义字段\n\
            picPath 插入图片的路径\n\
            copyUrl 备份的服务器路径\n\
            userName 传给wps要显示的OA用户名\n\
"
}

function onlineEditDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())
    var uploadFieldName = prompt("请输入文档上传到业务系统时自定义字段：", "自定义字段")
    _WpsInvoke([{
        "OnlineEditDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "uploadFieldName": uploadFieldName,
            "buttonGroups": "btnSaveAsFile,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark", //屏蔽功能按钮
            "userName": "东方不败"
        }
    }], true) // onlineEditDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['onlineEditDoc'] = {
    action: onlineEditDoc,
    code: _WpsInvoke.toString() + "\n\n" + onlineEditDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径在线打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        onlineEditDoc方法对应于OA助手dispatcher支持的方法名\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            uploadFieldName 文档上传到业务系统时自定义字段\n\
            buttonGroups 屏蔽的OA助手功能按钮\n\
            userName 传给wps要显示的OA用户名\n\
"
}

function openRevision() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传接口:")
    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "userName": '王五', //用户名
            "revisionCtrl": {
                "bOpenRevision": true,
                "bShowRevision": true
            }
        }
    }])
}

_wps['openRevision'] = {
    action: openRevision,
    code: _WpsInvoke.toString() + "\n\n" + openRevision.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，并打开修订功能，并显示修订\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        userName 用户名，设置当前编辑用户名\n\
        fileName 打开的文档路径\n\
        revisionCtrl 修订功能控制参数\n\
"
}

function closeRevision() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传接口:")
    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "userName": '王五', //用户名
            "revisionCtrl": {
                "bOpenRevision": false,
                "bShowRevision": false
            }
        }
    }])
}

_wps['closeRevision'] = {
    action: closeRevision,
    code: _WpsInvoke.toString() + "\n\n" + closeRevision.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，并关闭修订功能\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        userName 用户名，设置当前编辑用户名\n\
        fileName 打开的文档路径\n\
        revisionCtrl 修订功能控制参数\n\
"
}

function protectOpen() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传接口:")
    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "openType": { //文档打开方式
                // 文档保护类型，-1：不启用保护模式，0：只允许对现有内容进行修订，
                // 1：只允许添加批注，2：只允许修改窗体域(禁止拷贝功能)，3：只读
                "protectType": 3,
                "password": "123456"
            }
        }
    }])
}

_wps['protectOpen'] = {
    action: protectOpen,
    code: _WpsInvoke.toString() + "\n\n" + protectOpen.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开使用保护模式指定文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        uploadPath 保存文档上传接口\n\
        fileName 打开的文档路径\n\
        openType 文档打开方式控制参数 protectType：1：不启用保护模式，0：只允许对现有内容进行修订，\n\
        \t\t1：只允许添加批注，2：只允许修改窗体域(禁止拷贝功能)，3：只读 password为密码\n\
"
}

function openWithPassWd() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：")
    var docPassword = prompt("请输入文档打开密码:")
    var uploadPath = prompt("请输入文档上传接口:")
    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "docPassword": {
                "docPassword": docPassword // 文档密码
            }
        }
    }])
}

_wps['openWithPassWd'] = {
    action: openWithPassWd,
    code: _WpsInvoke.toString() + "\n\n" + openWithPassWd.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，使用指定密码打开指定加密文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        uploadPath 保存文档上传接口\n\
        fileName 打开的文档路径\n\
        docPassword 文档密码\n\
"
}

function insertRedHeader() {
    var filePath = prompt("请输入打开文件路径，如果为空则对活动文档套红：", GetDemoPath("样章.docx"))
    var templateURL = prompt("请输入红头模板路径（本地或是url）:", GetDemoPath("红头文件.docx"))
    if (filePath != '' && filePath != null) {
        _WpsInvoke([{
            "OnlineEditDoc": {
                "fileName": filePath,
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
                "buttonGroups": "btnSaveAsFile,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark" //屏蔽功能按钮
            }
        }])
    } else {
        _WpsInvoke([{
            "InsertRedHead": {
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content' //红头模板中填充正文的位置书签名
            }
        }])
    }
}

_wps['insertRedHeader'] = {
    action: insertRedHeader,
    code: _WpsInvoke.toString() + "\n\n" + insertRedHeader.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，然后使用指定红头模板对该文档进行套红头\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        fileName 打开的文档路径\n\
        insertFileUrl 指定的红头模板\n\
        bkInsertFile 红头模板中正文的位置书签名\n\
    InsertRedHead方法对应于OA助手dispatcher支持的方法名\n\
        insertFileUrl 指定的红头模板\n\
        bkInsertFile 红头模板中正文的位置书签名\n\
"
}

function fillTemplate() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章2.docx"))
    var templatePath = prompt("请输入需要填充的数据的请求地址:", document.location.protocol + "//" + document.location.host + "/getTemplateData")

    _WpsInvoke([{
        "OpenDoc": {
            "fileName": filePath,
            "templateDataUrl": templatePath
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['fillTemplate'] = {
    action: fillTemplate,
    code: _WpsInvoke.toString() + "\n\n" + fillTemplate.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，打开WPS文字后,将根据文档路径下载并打开对应的文档，\n\
    并自动从模板服务器获取模板数据并套用到文档中\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            fileName 打开的文档路径\n\
            templateDataUrl 模板的服务器路径\n\
"
}

function convertDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档转换后上传接口:", GetUploadPath())

    _WpsInvoke([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传接口
            "fileName": filePath,
            "suffix": ".pdf",
            "uploadWithAppendPath": "1" //与suffix配置使用，传入标志位即可
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['convertDoc'] = {
    action: convertDoc,
    code: _WpsInvoke.toString() + "\n\n" + convertDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档转换后上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径下载并打开对应的文档，转换完将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            suffix 转换类型\n\
            uploadWithAppendPath 保存时一并转换的目标格式\n\
"
}

function taskPaneBookMark() {
    var filePath = prompt("请输入打开带书签文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    _WpsInvoke([{
        "taskPaneBookMark": {
            "fileName": filePath,
            "userName": "东方不败"
        }
    }]) // taskPaneBookMark方法对应于OA助手dispatcher支持的方法名
}

_wps['taskPaneBookMark'] = {
    action: taskPaneBookMark,
    code: _WpsInvoke.toString() + "\n\n" + taskPaneBookMark.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，文档中的书签将会在OA助手的Taskpane中显示出来。\n\
    点击TaskPane中的书签，会自动跳转到书签所对应的文件中的位置，点击Taskpane按F2键可以看Taskpane中的html源码。\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    taskPaneBookMark 方法对应于OA助手dispatcher支持的方法名\n\
        fileName 文档的路径\n\
        userName 传给wps要显示的OA用户名\n\
"
}

function exitWPS() {
    _WpsInvoke([{
        "ExitWPS": {}
    }], true)
}

_wps['exitWPS'] = {
    action: exitWPS,
    code: _WpsInvoke.toString() + "\n\n" + exitWPS.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，关闭已打开的WPS\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议通知WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    ExitWPS 方法对应于OA助手dispatcher支持的方法名\n\
"
}

function getDocStatus() {
    _WpsInvoke([{
        "GetDocStatus": {}
    }],
        false)
}

_wps['getDocStatus'] = {
    action: getDocStatus,
    code: _WpsInvoke.toString() + "\n\n" + getDocStatus.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，获取活动文档的状态\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议通知WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    GetDocStatus 方法对应于OA助手dispatcher支持的方法名\n\
"
}

//----公文写作的相关方法 这些都必须是在有「公文写作」组件的版本中运行 Start--------
/**
 * 判断当前OS是否是Linux系统
 *
 * @returns
 */
function checkOSisLinux() {
    if (detectOS() == "Linux") {
        return true
    } else {
        alert("此方法仅在WPS Linux特定版本支持")
    }
}
/**
 * 新建一个使用公文写作打开的公文
 *
 */
function newOfficialDocument() {
    if (checkOSisLinux()) {
        _WpsInvoke([{
            "NewOfficialDocument": {
                "isOfficialDocument": true
            }
        }]) // NewOfficialDocument方法对应于OA助手dispatcher支持的方法名
    }
}

_wps['newOfficialDocument'] = {
    action: newOfficialDocument,
    code: _WpsInvoke.toString() + "\n\n" + newOfficialDocument.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，打开WPS公文写作后,新建一个公文\n\
\n\
  方法使用：\n\
    页面点击按钮， 通过wps客户端协议来启动WPS， 调用oaassist插件， 执行插件中的js函数NewOfficialDocument, 新建一个默认模板的公文\ n\
    funcs参数说明:\n\
        NewOfficialDocument方法对应于OA助手dispatcher支持的方法名\ n\
"
}

/**
 * 打开一个使用公文写作打开的公文
 */
function openOfficialDocument() {
    if (checkOSisLinux()) {
        var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("公文样章.wps"))
        var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())
        var uploadFieldName = prompt("请输入文档上传到业务系统时自定义字段：", "自定义字段")
        var backupPath = prompt("请输入文档备份路径:")
        _WpsInvoke([{
            "OpenDoc": {
                "uploadPath": uploadPath, // 保存文档上传接口
                "fileName": filePath,
                "uploadFieldName": uploadFieldName,
                "picPath": GetDemoPngPath(),
                "copyUrl": backupPath,
                "userName": "东方不败"
            }
        }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
    }
}
_wps['openOfficialDocument'] = {
    action: openOfficialDocument,
    code: _WpsInvoke.toString() + "\n\n" + openOfficialDocument.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径下载并打开对应的文档，保存将自动上传指定服务器地址\n\
\n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            uploadFieldName 文档上传到业务系统时自定义字段\n\
            picPath 插入图片的路径\n\
            copyUrl 备份的服务器路径\n\
            userName 传给wps要显示的OA用户名\n\
"
}
/**
 * 在线不落地打开一个使用公文写作打开的公文
 */
function onlineEditOfficialDocument() {
    if (checkOSisLinux()) {
        var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("公文样章.wps"))
        var uploadPath = prompt("请输入文档上传接口:", GetUploadPath())
        var uploadFieldName = prompt("请输入文档上传到业务系统时自定义字段：", "自定义字段")
        _WpsInvoke([{
            "OnlineEditDoc": {
                "uploadPath": uploadPath, // 保存文档上传接口
                "fileName": filePath,
                "uploadFieldName": uploadFieldName,
                "buttonGroups": "btnSaveAsFile,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark", //屏蔽功能按钮
                "userName": "东方不败"
            }
        }]) // onlineEditDoc方法对应于OA助手dispatcher支持的方法名
    }
}

_wps['onlineEditOfficialDocument'] = {
    action: onlineEditOfficialDocument,
    code: _WpsInvoke.toString() + "\n\n" + onlineEditOfficialDocument.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传接口，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径在线打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OnlineEditDoc方法对应于OA助手dispatcher支持的方法名\n\
            uploadPath 保存文档上传接口\n\
            fileName 打开的文档路径\n\
            uploadFieldName 文档上传到业务系统时自定义字段\n\
            buttonGroups 屏蔽的OA助手功能按钮\n\
            userName 传给wps要显示的OA用户名\n\
"
}

/** 
 * 这是HTML页面上的按钮赋予事件的实现，开发者无需关心，使用自己习惯的方式做开发即可
 */
window.onload = function () {
    var btns = document.getElementsByClassName("btn");

    for (var i = 0; i < btns.length; i++) {
        btns[i].onclick = function (event) {
            document.getElementById("blockFunc").style.visibility = "visible";
            var btn2 = document.getElementById("demoBtn");
            btn2.innerText = this.innerText;
            document.getElementById("codeDes").innerText = _wps[this.id].detail.toString()
            document.getElementById("code").innerText = _wps[this.id].code.toString()
            var onBtnAction = _wps[this.id].action

            // document.getElementById("demoBtn").onclick = onBtnAction //IE不支持箭头函数，改为通用写法
            document.getElementById("demoBtn").onclick = function () { //IE不支持箭头函数，改为通用写法
                //之下动作是做了对Node服务的判断和oem.ini的设置
                var xhr = new WpsInvoke.CreateXHR();
                xhr.onload = function () {
                    onBtnAction()
                }
                xhr.onerror = function () {
                    alert("请确认本地服务端(StartupServer.js)是启动状态")
                    return
                }
                xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
                xhr.send()
            }

            hljs.highlightBlock(document.getElementById("code"));
        }
    }
}
/**
 * 检查操作系统
 *
 * @returns Win10 | Win7 | WinVista | Win2003 | WinXP | Win2000 | Linux | Unix | Mac
 */
function detectOS() {
    var sUserAgent = navigator.userAgent;
    var isWin = (navigator.platform == "Win32") || (navigator.platform == "Windows");
    var isMac = (navigator.platform == "Mac68K") || (navigator.platform == "MacPPC") || (navigator.platform == "Macintosh") || (navigator.platform == "MacIntel");
    if (isMac) return "Mac";
    var isUnix = (navigator.platform == "X11") && !isWin && !isMac;
    if (isUnix) return "Unix";
    var isLinux = (String(navigator.platform).indexOf("Linux") > -1);
    if (isLinux) return "Linux";
    if (isWin) {
        var isWin2K = sUserAgent.indexOf("Windows NT 5.0") > -1 || sUserAgent.indexOf("Windows 2000") > -1;
        if (isWin2K) return "Win2000";
        var isWinXP = sUserAgent.indexOf("Windows NT 5.1") > -1 || sUserAgent.indexOf("Windows XP") > -1;
        if (isWinXP) return "WinXP";
        var isWin2003 = sUserAgent.indexOf("Windows NT 5.2") > -1 || sUserAgent.indexOf("Windows 2003") > -1;
        if (isWin2003) return "Win2003";
        var isWinVista = sUserAgent.indexOf("Windows NT 6.0") > -1 || sUserAgent.indexOf("Windows Vista") > -1;
        if (isWinVista) return "WinVista";
        var isWin7 = sUserAgent.indexOf("Windows NT 6.1") > -1 || sUserAgent.indexOf("Windows 7") > -1;
        if (isWin7) return "Win7";
        var isWin10 = sUserAgent.indexOf("Windows NT 6.1") > -1 || sUserAgent.indexOf("Windows 10") > -1;
        if (isWin10) return "Win10";
    }
    return "other";
}