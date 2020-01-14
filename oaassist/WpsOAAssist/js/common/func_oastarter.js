/**
 * 在这个js中，集中处理来自OA的传入参数
 * 
 */

/**
 * web页面调用WPS的方法入口
 *  * info参数结构
 * info:[
 *      {
 *       '方法名':'方法参数',需要执行的方法
 *     },
 *     ...
 *   ]
 * @param {*} info
 */
function dispatcher(info) {
    var funcs = info.funcs;

    //执行web页面传递的方法
    for (var index = 0; index < funcs.length; index++) {
        var func = funcs[index];
        for (var key in func) {
            if (key === "OpenDoc") { // OpenDoc 属于普通的打开文档的操作方式，文档落地操作
                OpenDoc(func[key]); //进入打开文档处理函数
            } else if (key === "OnlineEditDoc") { //在线方式打开文档，属于文档不落地的方式打开
                OnlineEditDoc(func[key]);
            } else if (key === "NewDoc") {
                OpenDoc(func[key]);
            } else if (key === "UseTemplate") {
                OpenDoc(func[key]);
            } else if (key === "InsertRedHead") {
                InsertRedHead(func[key]);
            } else if (key === "taskPaneBookMark"){
                taskPaneBookMark(func[key])
            }
        }
    }
}

/**
 * 
 * @param {*} params  OA端传入的参数
 */
function OnlineEditDoc(OaParams) {
    if (OaParams.fileName == "") {
        NewFile(OaParams);
    } else {
        //OA传来下载文件的URL地址，调用openFile 方法打开
        OpenOnLineFile(OaParams);
    }
}

///打开来自OA端传递来的文档
function OpenDoc(OaParams) {
    if (OaParams.fileName == "") {
        NewFile(OaParams);
    } else {
        //OA传来下载文件的URL地址，调用openFile 方法打开
        OpenFile(OaParams);
    }
}

function taskPaneBookMark(OaParams){
    let filePath = OaParams.fileName
    if (filePath == "")
        return
    OpenFile(OaParams);

    //创建taskpane，只创建一次
    let id = wps.PluginStorage.getItem(constStrEnum.taskpaneid)
    if (id){
        let tp = wps.GetTaskpane(id)
        tp.Visible = true
    }
    else{
        let url = getHtmlURL("taskpane.html");
        let tp =  wps.CreateTaskPane(url, "书签操作")
        if (tp){
            //tp.DockPosition = 
            tp.Width = 300
            tp.Visible = true
            wps.PluginStorage.setItem(constStrEnum.taskpaneid, tp.ID)
        }
    }
}