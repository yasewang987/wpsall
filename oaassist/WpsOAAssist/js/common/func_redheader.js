//套红头

/**
 *  套红头入口
 */
function OnInsertRedHeader(Doc) {
    if (!Doc) {
        return;
    }

    var height = 200;
    if (wps.PluginStorage.getItem("searchRedHeadPath")) {
        height = height + 50;
    }
    OnShowDialog("redhead.html", "OA助手", 400, height);
}