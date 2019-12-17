function selectTemplate() {
    var wpsApp = wps.WpsApplication();
    var obj = document.getElementById("templates");
    var index = obj.selectedIndex;
    if (index == -1) { //添加未选中数据时的异常处理
        alert("请先选择红头文件后再进行套红头！");
        return;
    }
    var redId = obj.options[index].getAttribute("value");
    var activeDoc = wpsApp.ActiveDocument;
    var base = wps.PluginStorage.getItem("getRedHeadPath") || OA_DOOR.getRedHeadPath;
    if (base == undefined) { //未配置则模拟服务端返回
        SetDocParamsValue(activeDoc, "insertFileUrl", getDemoTemplatePath());
        SetDocParamsValue(activeDoc, "bkInsertFile", "Content");
        InsertRedHeadDoc(activeDoc);
        window.opener = null;
        window.open('', '_self', '');
        window.close();
        return
    }
    var path = base + redId;
    SetDocParamsValue(activeDoc, "insertFileUrl", path);
    InsertRedHeadDoc(activeDoc);
    window.opener = null;
    window.open('', '_self', '');
    window.close();
}
//判断是否是word文档
function isWord(suffix) {
    var suffixArray = ["doc", "dot", "wps", "wpt", "docx", "docm", "dotm"];
    for (var f1 in suffixArray) {
        if (suffixArray[f1].indexOf(suffix) > -1) {
            return true;
        }
    }
    return false;
}

function getAllTemplelists() {
    var xmlhttp;
    if (window.XMLHttpRequest) { //IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp = new XMLHttpRequest();
    } else { // IE6, IE5
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    //上面的http请求对象的生成做了一个浏览器兼容性处理
    xmlhttp.onreadystatechange = function () {
        //当接受到响应时回调该方法
        if (xmlhttp.readyState == 4 && (xmlhttp.status == 200 || xmlhttp.status == 0)) {
            var text = xmlhttp.responseText; //使用接口返回内容，响应内容
            var resultJson = JSON.parse(text) //将json字符串转换成对象    
            for (var i = 0; i < resultJson.length; i++) {
                var element = resultJson[i]
                var myOption = document.createElement("option"); //动态创建option标签
                var suffix = element.tempName.split('.')[1];
                if (isWord(suffix)) {
                    myOption.value = element.tempId; //红头文档id
                    myOption.text = element.tempName; //红头文档名称
                    templates.add(myOption);
                }
            }
        }
    }

    var redHeadsPath = wps.PluginStorage.getItem("redHeadsPath") || OA_DOOR.redHeadsPath
    if (redHeadsPath == undefined) { //未配置则模拟服务端返回
        var strData =
            '[{"tempId":23,"tempName":"广东省xx局.docx"},{"tempId":24,"tempName":"广东省xx局 - 副本.docx"},{"tempId":25,"tempName":"红头.docx"}]'
        resultJson = JSON.parse(strData);
        for (var i = 0; i < resultJson.length; i++) {
            var element = resultJson[i]
            var myOption = document.createElement("option"); //动态创建option标签
            var suffix = element.tempName.split('.')[1];
            if (isWord(suffix)) {
                myOption.value = element.tempId; //红头文档id
                myOption.text = element.tempName; //红头文档名称
                templates.add(myOption);
            }
        }
        document.getElementById("search").style.display = "none";
        return
    }
    xmlhttp.open("POST", redHeadsPath, true); //以POST方式请求该接口
    xmlhttp.setRequestHeader("Content-type",
        "application/x-www-form-urlencoded;charset=UTF-16LE"); //添加Content-type
    xmlhttp.send(); //发送请求参数间用&分割
    if (!wps.PluginStorage.getItem("searchRedHeadPath")) {
        document.getElementById("search").style.display = "none";
    }
}

function search() {
    var xmlhttp;
    if (window.XMLHttpRequest) { //IE7+, Firefox, Chrome, Opera, Safari
        xmlhttp = new XMLHttpRequest();
    } else { // IE6, IE5
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    //上面的http请求对象的生成做了一个浏览器兼容性处理
    xmlhttp.onreadystatechange = function () {
        //当接受到响应时回调该方法
        if (xmlhttp.readyState == 4 && (xmlhttp.status == 200 || xmlhttp.status == 0)) {
            var text = xmlhttp.responseText; //使用接口返回内容，响应内容
            var resultJson = JSON.parse(text) //将json字符串转换成对象
            templates.options.length = 0;
            for (var i = 0; i < resultJson.length; i++) {
                var element = resultJson[i]
                var myOption = document.createElement("option"); //动态创建option标签
                myOption.value = element.tempId; //红头文档id
                myOption.text = element.tempName; //红头文档名称
                templates.add(myOption);
            }
        }
    }

    var searchPath = wps.PluginStorage.getItem("searchRedHeadPath") || OA_DOOR.redHeadsPath

    var totalPath = searchPath + "?content=" + document.getElementById("content").value;

    xmlhttp.open("get", totalPath, true); //以POST方式请求该接口
    xmlhttp.setRequestHeader("Content-type",
        "application/x-www-form-urlencoded;charset=UTF-16LE"); //添加Content-type
    xmlhttp.send(); //发送请求参数间用&分割
}