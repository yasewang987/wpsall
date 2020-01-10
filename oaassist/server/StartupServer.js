const express = require('express');
const fs = require('fs');
const path = require('path');
var urlencode = require('urlencode');
const formidable = require('formidable')
var ini = require('ini')
var regedit = require('regedit')

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
	console.log(getNow() + filePath)
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
app.use("/WpsSetup", (request, response) => {
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});

//wps安装包是否正确的检测
app.use("/OAAssistDeploy", (request, response) => {
	response.writeHead(200, "OK", { "Content-Type": "text/html; charset=utf-8" })
	response.end("成功");
});

app.use("/Download/:fileName", function (request, response) {
	var fileName = request.params.fileName;
	var filePath = path.join(__dirname, './wwwroot/file');
	filePath = path.join(filePath, fileName);
	var stats = fs.statSync(filePath);
	if (stats.isFile()) {
		let name = urlencode(fileName, "utf-8");
		response.set({
			'Content-Type': 'application/octet-stream',
			//'Content-Disposition': "attachment; filename* = UTF-8''" + name,
			'Content-Disposition': "attachment; filename=" + name,
			'Content-Length': stats.size
		});
		fs.createReadStream(filePath).pipe(response);
		console.log(getNow() + "下载文件接被调用，文件路径：" + filePath)
	} else {
		response.writeHead(200, "Failed", { "Content-Type": "text/html; charset=utf-8" });
		response.end("文件不存在");
	}
});
app.post("/Upload", function (request, response) {
	const form = new formidable.IncomingForm();
	form.encoding = 'utf-8';
	form.uploadDir = "wwwroot/tmp";
	form.parse(request, function (error, fields, files) {
		for (let key in files) {
			let file = files[key]
			// 过滤空文件
			if (file.size == 0 && file.name == '') continue

			let oldPath = file.path
			let	newPath = 'wwwroot/file/' + file.name

			fs.rename(oldPath, newPath, function (error) {
				console.log(getNow() + "上传文件成功，路径：" + newPath)
			})
		}
		response.writeHead(200, {
			"Content-Type": "text/html;charset=utf-8"
		})
		response.end("OK");
	})
});

function getNow(){
	let nowDate = new Date()
	let year = nowDate.getFullYear()
	let month = nowDate.getMonth()+1
	let day = nowDate.getDate()
	let hour = nowDate.getHours()
	let minute = nowDate.getMinutes()
	let second = nowDate.getSeconds()
 	return year+'年'+month+'月'+day+'日 '+hour+':'+minute+':'+second + "  "
}

function configOem(callback) {
	let oemPath;
	var key = 'HKCR\\KET.Sheet.12\\shell\\open\\command'
	var path = regedit.list(key, function (error, e) {
		try {
			var val = e[key].values[''].value;
			var pos = val.indexOf("et.exe");
			if (pos < 0){
				return callback({status:1,msg:"wps安装异常，请确认有没有正确的安装wps2019企业版！"})
			}
			var path = val.substring(1, pos) + 'cfgs\\oem.ini';
			oemPath = path
			var config = ini.parse(fs.readFileSync(path, 'utf-8'))
			var sup = config.support || config.Support;
			var ser = config.server || config.Server;
			var needUpdate = false;
			if (!sup || !sup.JsApiPlugin || !sup.JsApiShowWebDebugger)
				needUpdate = true;
			if (!ser || !ser.JSPluginsServer || ser.JSPluginsServer != "http://127.0.0.1:3888/jsplugins.xml")
				needUpdate = true;
			if (!sup) {
				sup = {}
				config.Support = sup
			}
			if (!ser) {
				ser = {}
				config.Server = ser
			}
			sup.JsApiPlugin = true
			sup.JsApiShowWebDebugger = true
			ser.JSPluginsServer = "http://127.0.0.1:3888/jsplugins.xml"

			if (needUpdate)
				fs.writeFileSync(path, ini.stringify(config))
		} catch (e) {
			oemResult = "配置" + oemPath + "失败，请尝试以管理员重新运行！！";
			console.log(oemResult)
			return callback({ status: 1, msg: oemResult })
		}
		callback({ status: 0, msg: "wps安装正常，" + oemPath + "文件设置正常。" })
	})
}

app.use("/WpsSetupTest", function (request, response) {
	configOem(function (res) {
		response.writeHead(200, res.status, {
			"Content-Type": "text/html;charset=utf-8"
		})
		response.end(res.msg);
	});
});

app.use(express.static(path.join(__dirname, "wwwroot"))); //wwwroot代表http服务器根目录
app.use('/plugin/et', express.static('../EtOAAssist'));
app.use('/plugin/wps', express.static('../WpsOAAssist'));
app.use('/plugin/wpp', express.static('../WppOAAssist'));

var server = app.listen(3888, function() {
	console.log(getNow() + "启动本地web服务(http://127.0.0.1:3888)成功！")
});

server.on('error', (e) => {
	if (e.code === 'EADDRINUSE') {
		console.log('地址正被使用，重试中...');
		setTimeout(() => {
			server.close();
			server.listen(3888);
		}, 2000);
	}
});