window.onload = function () {
    document.getElementById('btnMakeQRCode').onclick = function () {
        document.getElementById("qrcode").innerHTML = '';
        var text = document.getElementById('text').value;
        var qrcode = new QRCode(document.getElementById("qrcode"), {
            width: 100, //设置宽高
            height: 100
        });
        qrcode.makeCode(text);
    }

    document.getElementById('btnInsert').onclick = function () {
        let l_canvasImg = $("#qrcode").find("canvas")[0];
        let l_DataURL = l_canvasImg.toDataURL("image/png");
        let l_shapeQR = wps.WpsApplication().ActiveDocument.Shapes.AddBase64Picture(l_DataURL);
        l_shapeQR.Visible = true;
        l_shapeQR.Select();
    }

    function makeCode() {
        let elText = document.getElementById("text");
        if (!elText.value) {
            alert("请输入二维码文字");
            elText.focus();
            return;
        }
        qrcode.makeCode(elText.value);
    }

    $("#text").
    on("blur", function () {
        makeCode();
    }).
    on("keydown", function (e) {
        if (e.keyCode == 13) {
            makeCode();
        }
    });
}