const express = require('express');
const fs = require('fs');
const path = require('path');
var urlencode = require('urlencode');
const formidable = require('formidable')

const app = express()

//设置允许跨域访问该服务.
app.all('*', function (req, res, next) {
	res.header('Access-Control-Allow-Origin', '*');
	//Access-Control-Allow-Headers ,可根据浏览器的F12查看,把对应的粘贴在这里就行
	//res.header('Access-Control-Allow-Headers', 'Content-Type');
	//res.header('Access-Control-Allow-Methods', '*');
	//res.header('Content-Type', 'text/html;charset=utf-8');
	next();
  });

app.use("/FileList", function (request, response) {
	var filePath = path.join(__dirname, './wwwroot/file');
	console.log(filePath)
	fs.readdir(filePath, function (err, results) {
		if (err) {
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end("没有找到file文件夹");
			return;
		}
		if (results.length > 0) {
			var files = [];
			results.forEach(function (file) {
				if (fs.statSync(path.join(filePath, file)).isFile()) {
					files.push(file);
				}
			})
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end(files.toString());
		} else {
			response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" });
			response.end("当前目录下没有文件");
		}
	});
});

//wps安装包是否正确的检测
app.use("/WpsSetup", (request, response)=>{
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});

//wps安装包是否正确的检测
app.use("/OAAssistDeploy", (request, response)=>{
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});

app.use("/Download/:fileName", function (request, response) {
	var fileName = request.params.fileName;
	var filePath = path.join(__dirname, './wwwroot/file');
	filePath = path.join(filePath, fileName);
	console.log(filePath)
	var stats = fs.statSync(filePath);
	if (stats.isFile()) {
		let name = urlencode(fileName, "utf-8");
		response.set({
			'Content-Type': 'application/octet-stream',
			'Content-Disposition': "attachment; filename* = UTF-8''" + name,
			'Content-Length': stats.size
		});
		fs.createReadStream(filePath).pipe(response);
	} else {
		response.writeHead(200, "Failed", { "Content-Type": "text/html; charset=utf-8" });
		response.end("文件不存在");
	}
});
app.post("/Upload", function (request, response) {
	const form = new formidable.IncomingForm();
	form.encoding = 'utf-8';
	form.uploadDir = "tmp";
	form.parse(request, function (error, fields, files) {
		for (let key in files) {
			let file = files[key]
			// 过滤空文件
			if (file.size == 0 && file.name == '') continue

			let oldPath = file.path,
				newPath = 'wwwroot/file/' + file.name

			fs.rename(oldPath, newPath, function (error) {
				console.info(newPath)
			})
		}
		response.writeHead(200, {
			"Content-Type": "text/html;charset=utf-8"
		})
		response.end("OK");
	})
});

function getLocalIP() {
    const os = require('os');
    const osType = os.type(); //系统类型
    const netInfo = os.networkInterfaces(); //网络信息
    let ip = '';
    if (osType === 'Windows_NT') { 
        for (let dev in netInfo) {
        	//win7的网络信息中显示为本地连接，win10显示为以太网
            if (dev === '本地连接' || dev === '以太网') {
                for (let j = 0; j < netInfo[dev].length; j++) {
                    if (netInfo[dev][j].family === 'IPv4') {
                        ip = netInfo[dev][j].address;
                        break;
                    }
                }
            }
        }

    } else if (osType === 'Linux') {
        ip = netInfo.eth0[0].address;
    }

    return ip;
}

app.use(express.static(path.join(__dirname,"wwwroot"))); //wwwroot代表http服务器根目录
app.use('/plugin/et', express.static('../EtOAAssist'));
app.use('/plugin/wps', express.static('../WpsOAAssist'));
app.use('/plugin/wpp', express.static('../WppOAAssist'));

var server = app.listen(3888, function () {
	const host = server.address().address
	const port = server.address().port
	console.log("启动本地web服务(http://127.0.0.1:3888)成功！")
	/*console.log("监听 http://%s:%s", "127.0.0.1", port)
	console.log("监听 http://%s:%s", "localhost", port)
	console.log("监听 http://%s:%s", getLocalIP(), port)
	console.log("获取文件列表地址为 http://%s:%s/FileList", getLocalIP(), port)
	console.log("下载地址为 http://%s:%s/Download/*文件名", getLocalIP(), port)
	console.log("上传地址为 http://%s:%s/Upload", getLocalIP(), port)*/

	
});