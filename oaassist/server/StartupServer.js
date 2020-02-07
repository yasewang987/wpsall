const express = require('express');
const fs = require('fs');
const path = require('path');
var urlencode = require('urlencode');
const formidable = require('formidable')
var ini = require('ini')
var regedit = require('regedit')
const os = require('os');
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

//获取file目录下文件列表
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

//文件下载
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

//文件上传
app.post("/Upload", function (request, response) {
	const form = new formidable.IncomingForm();
	var uploadDir = path.join(__dirname, './wwwroot/uploaded/');
	form.encoding = 'utf-8';
	form.uploadDir = uploadDir;
	console.log(getNow() + "上传文件夹地址是：" + uploadDir);
	//判断上传文件夹地址是否存在，如果不存在就创建
	if (!fs.existsSync(form.uploadDir)) {
		fs.mkdirSync(form.uploadDir);
	}
	form.parse(request, function (error, fields, files) {
		for (let key in files) {
			let file = files[key]
			// 过滤空文件
			if (file.size == 0 && file.name == '') continue

			let oldPath = file.path
			let newPath = uploadDir + file.name

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

//获取当前时间
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

function configOemFileInner(oemPath, callback) {
	var config = ini.parse(fs.readFileSync(oemPath, 'utf-8'))
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
		fs.writeFileSync(oemPath, ini.stringify(config))

	callback({ status: 0, msg: "wps安装正常，" + oemPath + "文件设置正常。" })
}

function configOem(callback) {
	let oemPath;
	try {
		if (os.platform() == 'win32') {
			var key = 'HKCR\\KET.Sheet.12\\shell\\open\\command'
			regedit.list(key, function (error, e) {
				if (typeof (e) == "undefined" || e == null) {
					return callback({
						status: 1,
						msg: "WPS未安装，请安装WPS2019企业版。"
					})
				}
				var val = e[key].values[''].value;
				var pos = val.indexOf("et.exe");
				if (pos < 0) {
					return callback({ status: 1, msg: "wps安装异常，请确认有没有正确的安装wps2019企业版！" })
				}
				oemPath = val.substring(1, pos) + 'cfgs\\oem.ini';
				configOemFileInner(oemPath, callback)
			})
		} else {
			oemPath = "/opt/kingsoft/wps-office/office6/cfgs/oem.ini";
			configOemFileInner(oemPath, callback);
		}
	} catch (e) {
		oemResult = "配置" + oemPath + "失败，请尝试以管理员重新运行！！";
		console.log(oemResult)
		return callback({ status: 1, msg: oemResult })
	}
}

app.use("/WpsSetupTest", function (request, response) {
	configOem(function (res) {
		response.writeHead(200, res.status, {
			"Content-Type": "text/html;charset=utf-8"
		});
		response.write('<head><meta charset="utf-8"/></head>');
		response.write("<br/>当前检测时间为： " + getNow() + "<br/>");
		response.end(res.msg);
	});
});

//支持jsplugins.xml中，在线模式下，WPS加载项的请求地址
app.use(express.static(path.join(__dirname, "wwwroot"))); //wwwroot代表http服务器根目录
app.use('/plugin/et', express.static(path.join(__dirname, "../EtOAAssist")));
app.use('/plugin/wps', express.static(path.join(__dirname, "../WpsOAAssist")));
app.use('/plugin/wpp', express.static(path.join(__dirname, "../WppOAAssist")));

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

//获取服务端IP地址
function getServerIPAdress() {
	var interfaces = require('os').networkInterfaces();
	for (var devName in interfaces) {
		var iface = interfaces[devName];
		for (var i = 0; i < iface.length; i++) {
			var alias = iface[i];
			if (alias.family === 'IPv4' && alias.address !== '127.0.0.1' && !alias.internal) {
				return alias.address;
			}
		}
	}
}