if (typeof (wps.Enum) == "undefined") { // 如果没有内置枚举值
    wps.Enum = WPS_Enum;
}
/**
 * WPS常用的API枚举值，具体参与API文档
 */
var WPS_Enum = { 
    wdDoNotSaveChanges: 0,
    wdFormatPDF: 17,
    wdFormatOpenDocumentText: 23,
    wdFieldFormTextInput: 70,
    wdAlertsNone: 0,
    wdDialogFilePageSetup: 178,
    wdDialogFilePrint: 88,
    wdRelativeHorizontalPositionPage: 1,
    wdGoToPage: 1,
    wdPropertyPages: 14,
    wdRDIComments: 1,
    wdDialogInsertDateTime: 165,
    msoCTPDockPositionLeft: 0,
    msoCTPDockPositionRight: 2
}
/**
 * WPS图片版式枚举
 */
var WPS_Enum_WdWrapType={
    wdWrapInline: 7, //将形状嵌入到文字中。
    wdWrapNone: 3, //将形状放在文字前面。 请参阅 wdWrapFront。
    wdWrapSquare: 0, //使文字环绕形状。 行在形状的另一侧延续。
    wdWrapThrough: 2, //使文字环绕形状。
    wdWrapTight: 1, //使文字紧密地环绕形状。
    wdWrapTopBottom: 4, //将文字放在形状的上方和下方。
    wdWrapBehind: 5, //将形状放在文字后面。
    wdWrapFront: 6 //将形状放在文字前面。
}

/**
 * WPS加载项自定义的枚举值
 */
var constStrEnum = {
    AllowOADocReOpen: "AllowOADocReOpen",
    AutoSaveToServerTime: "AutoSaveToServerTime",
    bkInsertFile: "bkInsertFile",
    buttonGroups: "buttonGroups",
    CanSaveAs: "CanSaveAs",
    copyUrl: "copyUrl",
    DefaultUploadFieldName: "DefaultUploadFieldName",
    disableBtns: "disableBtns",
    insertFileUrl: "insertFileUrl",
    IsInCurrOADocOpen: "IsInCurrOADocOpen",
    IsInCurrOADocSaveAs: "IsInCurrOADocSaveAs",
    isOA: "isOA",
    notifyUrl: "notifyUrl",
    OADocCanSaveAs: "OADocCanSaveAs",
    OADocLandMode: "OADocLandMode",
    OADocUserSave: "OADocUserSave",
    openType: "openType",
    picPath: "picPath",
    picHeight: "picHeight",
    picWidth: "picWidth",
    redFileElement: "redFileElement",
    revisionCtrl: "revisionCtrl",
    ShowOATabDocActive: "ShowOATabDocActive",
    SourcePath: "SourcePath",
    suffix: "suffix",
    templateDataUrl: "templateDataUrl",
    TempTimerID: "TempTimerID",
    uploadPath: "uploadPath",
    uploadFieldName: "uploadFieldName",
    uploadFileName: "uploadFileName",
    uploadAppendPath: "uploadAppendPath",
    uploadWithAppendPath: "uploadWithAppendPath",
    userName: "userName",
    WPSInitUserName: "WPSInitUserName",
    taskpaneid: "taskpaneid"
}