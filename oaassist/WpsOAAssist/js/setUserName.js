function changeUserName() {
    var l_Name = $("#txtUserName").val();
    if ($.trim(l_Name) == "") {
        alert("请输入有效的用户名称");
        return;
    }
    l_Name = $.trim(l_Name);
    wps.PluginStorage.setItem("WPSInitUserName", l_Name);
    wps.WpsApplication().ActiveDocument.Activate();
    OnWindowActivate();
    window.close();
}