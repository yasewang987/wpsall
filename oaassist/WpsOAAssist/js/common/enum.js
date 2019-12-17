var WPS_Enum = { //wps api常用枚举值 具体参阅API文档
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
}

if (typeof (wps.Enum) == "undefined") { // 如果没有内置枚举值
    wps.Enum = WPS_Enum;
}