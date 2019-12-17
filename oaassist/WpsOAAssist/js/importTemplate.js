/**
 * 导入公文模板，并替换当前文档全部内容
 * @param templateURL  模板路径
 */
function importTemlateFile(templateURL) {
    var wpsApp = wps.WpsApplication();
    var activeDoc = wpsApp.ActiveDocument;
    if (!activeDoc) {
        alert("文档不存在");
        return;
    }
    var selection = wpsApp.ActiveWindow.Selection;
    selection.WholeStory(); //选取全文
    selection.Delete(); // 删除选中内容
    selection.InsertFile(templateURL);
    if (activeDoc.Revisions.Count > 0) { // 文档或区域中的修订
        activeDoc.AcceptAllRevisions(); // 接受对指定文档的所有修订
    }
}

// 获取选中项，拼接模板Url进行导入模板
function OnImportTemplate() {
    var p_Doc = wps.WpsApplication().ActiveDocument;
    var templatePath = GetDocParamsValue(p_Doc, "templatePath");
    if (templatePath == "") {
        templatePath = OA_DOOR.templateBaseURL;
    }
    var templateId = vue.templateItem;
    var templateURL;
    if (templatePath == undefined)
        templateURL = getDemoTemplatePath();
    else
        templateURL = templatePath + templateId;

    importTemlateFile(templateURL);
    window.opener = null;
    window.open('', '_self', '');
    window.close();
}


function getAllTemplate() {
    var p_Doc = wps.WpsApplication().ActiveDocument;
    var templateData = GetDocParamsValue(p_Doc, "templateData");
    if (templateData == "") {
        templateData = OA_DOOR.templateDataUrl;
    }
    if (templateData == undefined) { //未配置的话，此处模拟网络请求的返回
        var strData =
            '[{"tempId":60,"tempName":"固定IP地址申请表2.xls"},{"tempId":65,"tempName":"企业秘模板.doc"},{"tempId":72,"tempName":"1呈批件（单一内设机构）.docx"},{"tempId":73,"tempName":"1呈批件（单一内设机构）2.docx"},{"tempId":74,"tempName":"企业秘模板.doc"},{"tempId":76,"tempName":"7电话记录11.docx"},{"tempId":77,"tempName":"4呈批件（两个部门三个内设机构）11.docx"},{"tempId":78,"tempName":"1呈批件（单一内设机构）11.docx"},{"tempId":79,"tempName":"OA模板：xx市高级人民法院排版用模版(无审判长).doc"},{"tempId":80,"tempName":"OA模板：xx市高级人民法院排版用模版(有审判长).doc"},{"tempId":81,"tempName":"OA模板1.doc"},{"tempId":82,"tempName":"测试插入速度空表格.xlsx"},{"tempId":97,"tempName":"模板.docx"},{"tempId":98,"tempName":"测试模板.docx"}]';
        vue.templates = JSON.parse(strData);
        console.log("数据：" + vue.templates);
        return;
    }
    $.ajax({
        url: templateData,
        async: false,
        method: "post",
        dataType: 'json',
        success: function (res) {
            vue.templates = res;
            console.log("模板列表数据：" + JSON.stringify(res));
        },
        error: function (res) {
            alert("获取响应失败");
            vue.templates = {}
        }
    });
}

function cancel() { // 取消按钮
    window.close();
}

var vue = new Vue({
    el: "#template",
    data: {
        templateItem: "",
        templates: {}
    },
    methods: {},
    created: function () {
        setTimeout(getAllTemplate, 0);
    }
});