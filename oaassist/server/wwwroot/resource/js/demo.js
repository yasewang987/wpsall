// 切换到相应 tab. 0: wps  1: wpp  2: et
function SwitchTab(crtTabIndex) {
    var iframe_wps = document.getElementById("iframe_wps");
    var iframe_wpp = document.getElementById("iframe_wpp");
    var iframe_et = document.getElementById("iframe_et");

    if (iframe_wps && iframe_wpp && iframe_et) {
        switch (crtTabIndex) {
            case 0:
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wps.setAttribute("height", "100%");
                break;
            case 1:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "100%");
                break;
            case 2:
                iframe_wps.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "100%");
                break;
            default:
                iframe_wps.setAttribute("height", "0px");
                iframe_wpp.setAttribute("height", "0px");
                iframe_et.setAttribute("height", "0px");
                break;
        }
    }
}

function appChange() {
    var obj = document.getElementById("app");
    SwitchTab(obj.selectedIndex);
}

window.onload = function () {
    // 默认切换到 wps
    SwitchTab(0);
    var app = document.getElementById("app");
    var opts = app.getElementsByTagName("option"); //得到数组option
    opts[0].selected = true;
}