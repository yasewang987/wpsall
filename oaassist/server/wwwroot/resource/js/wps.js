var _wps = {}

function _WpsStartUp(funcs) {
    var info = {};

    info.funcs = funcs;
    WpsStartUp.StartUp(WpsStartUp.ClientType.wps, // 组件类型
        "WpsOAAssist", // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (result.status)
                alert(result.message)
        })
}

function newDoc() {
    _WpsStartUp([{
        "NewDoc": {}
    }]) // NewDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['newDoc'] = {
    action: newDoc,
    code: _WpsStartUp.toString() + "\n\n" + newDoc.toString(),
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
    var url = document.location.toString();
    url = decodeURI(url);
    if (url.indexOf("/") != -1) {
        url = url.substring(0, url.lastIndexOf("/"));
    }
    if (url.length !== 0)
        url = url.concat("/" + fileName);

    return url;
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
    var uploadPath = prompt("请输入文档上传路径:")
    var backupPath = prompt("请输入文档备份路径:")

    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "picPath": GetDemoPngPath(),
            "copyUrl": backupPath,
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['openDoc'] = {
    action: openDoc,
    code: _WpsStartUp.toString() + "\n\n" + openDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传路径，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径下载并打开对应的文档，保存将自动上传指定服务器地址\n\
\n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID，OA助手用以标记文档的信息，以区分其他文档\n\
            uploadPath 保存文档上传路径\n\
            fileName 打开的文档路径\n\
            picPath 插入图片的路径\n\
            copyUrl 备份的服务器路径\n\
"
}

function onlineEditDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：")
    var uploadPath = prompt("请输入文档上传路径:")

    _WpsStartUp([{
        "OnlineEditDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "buttonGroups": "btnSaveAsFile,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark", //屏蔽功能按钮
        }
    }]) // onlineEditDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['onlineEditDoc'] = {
    action: onlineEditDoc,
    code: _WpsStartUp.toString() + "\n\n" + onlineEditDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传路径，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径在线打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        onlineEditDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID\n\
            uploadPath 保存文档上传路径\n\
            fileName 打开的文档路径\n\
            buttonGroups 屏蔽的OA助手功能按钮\n\
"
}

function openRevision() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传路径:")
    _WpsStartUp([{
        "OpenDoc": {
            "uploadPath": uploadPath, // 保存文档上传路径
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
    code: _WpsStartUp.toString() + "\n\n" + openRevision.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，并打开修订功能，并显示修订\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        docId 文档ID\n\
        userName 用户名，设置当前编辑用户名\n\
        fileName 打开的文档路径\n\
        revisionCtrl 修订功能控制参数\n\
"
}

function closeRevision() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传路径:")
    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
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
    code: _WpsStartUp.toString() + "\n\n" + closeRevision.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，并关闭修订功能\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        docId 文档ID\n\
        userName 用户名，设置当前编辑用户名\n\
        fileName 打开的文档路径\n\
        revisionCtrl 修订功能控制参数\n\
"
}

function protectOpen() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档上传路径:")
    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
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
    code: _WpsStartUp.toString() + "\n\n" + protectOpen.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开使用保护模式指定文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        uploadPath 保存文档上传路径\n\
        docId 文档ID\n\
        fileName 打开的文档路径\n\
        openType 文档打开方式控制参数 protectType：1：不启用保护模式，0：只允许对现有内容进行修订，\n\
        \t\t1：只允许添加批注，2：只允许修改窗体域(禁止拷贝功能)，3：只读 password为密码\n\
"
}

function openWithPassWd() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：")
    var docPassword = prompt("请输入文档打开密码:")
    var uploadPath = prompt("请输入文档上传路径:")
    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "docPassword": {
                "docPassword": docPassword // 文档密码
            }
        }
    }])
}

_wps['openWithPassWd'] = {
    action: openWithPassWd,
    code: _WpsStartUp.toString() + "\n\n" + openWithPassWd.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，使用指定密码打开指定加密文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        uploadPath 保存文档上传路径\n\
        docId 文档ID\n\
        fileName 打开的文档路径\n\
        docPassword 文档密码\n\
"
}

function GetDemoRedHeadPath() {
    var url = document.location.toString();
    url = decodeURI(url);
    if (url.indexOf("/") != -1) {
        url = url.substring(0, url.lastIndexOf("/"));
    }
    if (url.length !== 0)
        url = url.concat("/红头文件.docx");

    if (url.startsWith("file:///"))
        url = url.substr("file:///".length);
    return url;
}

function insertRedHeader() {
    var filePath = prompt("请输入打开文件路径，如果为空则对活动文档套红：")
    var templateURL = prompt("请输入红头模板路径（本地或是url）:", GetDemoRedHeadPath())
    if (filePath != '' && filePath != null) {
        _WpsStartUp([{
            "OpenDoc": {
                "docId": "123", // 文档ID
                "fileName": filePath,
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
            }
        }])
    } else {
        _WpsStartUp([{
            "InsertRedHead": {
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
            }
        }])
    }
}

_wps['insertRedHeader'] = {
    action: insertRedHeader,
    code: _WpsStartUp.toString() + "\n\n" + insertRedHeader.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，然后使用指定红头模板对该文档进行套红头\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        docId 文档ID\n\
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
    var templatePath = prompt("请输入模板地址:")

    _WpsStartUp([{
        "OpenDoc": {
            "docId": "c2de1fcd1d3e4ac0b3cda1392c36c9", // 文档ID
            "fileName": filePath,
            "templateDataUrl": templatePath,
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['fillTemplate'] = {
    action: fillTemplate,
    code: _WpsStartUp.toString() + "\n\n" + fillTemplate.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传路径，打开WPS文字后,将根据文档路径下载并打开对应的文档，\n\
    并自动从模板服务器获取模板数据并套用到文档中\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID，OA助手用以标记文档的信息，以区分其他文档\n\
            fileName 打开的文档路径\n\
            templateDataUrl 模板的服务器路径\n\
"
}

function convertDoc() {
    var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    var uploadPath = prompt("请输入文档转换后上传路径:")

    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadAppendPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "suffix": ".pdf ",
        }
    }]) // OpenDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['convertDoc'] = {
    action: convertDoc,
    code: _WpsStartUp.toString() + "\n\n" + convertDoc.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档转换后上传路径，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径下载并打开对应的文档，转换完将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID，OA助手用以标记文档的信息，以区分其他文档\n\
            uploadAppendPath 保存文档上传路径\n\
            fileName 打开的文档路径\n\
            suffix 转换类型\n\
"
}

window.onload = function () {
    var btns = document.getElementsByClassName("btn");

    for (var i = 0; i < btns.length; i++) {
        btns[i].onclick = function (event) {
            document.getElementById("blockFunc").style.visibility = "visible";
            var btn2 = document.getElementById("demoBtn");
            btn2.innerText = this.innerText;
            document.getElementById("codeDes").innerText = _wps[this.id].detail.toString()
            document.getElementById("code").innerText = _wps[this.id].code.toString()
            document.getElementById("demoBtn").onclick = _wps[this.id].action;

            hljs.highlightBlock(document.getElementById("code"));
        }
    }
}