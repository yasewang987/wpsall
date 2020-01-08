;
(function (global, factory) {
	typeof exports === 'object' && typeof module !== 'undefined' ?
		module.exports = factory(global) :
		typeof define === 'function' && define.amd ?
		define(factory) : factory(global)
}((
	typeof self !== 'undefined' ? self :
	typeof window !== 'undefined' ? window :
	typeof global !== 'undefined' ? global :
	this
), function (global) {
	global = global || {};
	var tryCount = 0;
	var bFinished = true;
	var timeout = 2000;

	function getHttpObj() {
		var httpobj = null;
		if (IEVersion() < 10) {
			try {
				httpobj = new XDomainRequest();
			} catch (e1) {
				httpobj = new XMLHttpRequest();
			}
		} else {
			httpobj = new XMLHttpRequest();
		}
		return httpobj;
	}

	function startWps(clientType, t, callback) {
		if (!bFinished) {
			if (callback)
				callback({
					status: 1,
					message: "上一次请求没有完成"
				});
			return;
		}
		tryCount = 0;
		bFinished = false;

		function startWpsInnder() {
			++tryCount;
			if (tryCount > 4) {
				if (bFinished)
					return;
				bFinished = true;
				timeout = 2000;
				if (callback)
					callback({
						status: 2,
						message: "请允许浏览器打开WPS Office"
					});
				return;
			}
			var xmlReq = getHttpObj();
			xmlReq.open('POST', "http://localhost:58890/" + clientType + "/runParams");
			xmlReq.onload = function (res) {
				tryCount = 4;
				bFinished = true;
				timeout = 2000;
				if (callback)
					callback({
						status: 0
					});
			}
			xmlReq.ontimeout = xmlReq.onerror = function (res) {
				xmlReq.bTimeout = true;
				if (tryCount == 1) { //打开wps并传参
					window.location.href = "ksoWPSCloudSvr://start=RelayHttpServer" //是否启动wps弹框
					timeout = 3000;
				}
				setTimeout(startWpsInnder, 1000);
			}
			if (IEVersion() < 10) {
				timeout = 3000;
				xmlReq.onreadystatechange = function () {
					if (xmlReq.readyState != 4)
						return;
					if (xmlReq.bTimeout) {
						return;
					}
					if (xmlReq.status === 200)
						xmlReq.onload();
					else
						xmlReq.onerror();
				}
			}
			xmlReq.timeout = timeout;
			xmlReq.send("ksowebstartup" + clientType + "://" + t)
		}
		startWpsInnder();
		return;
	}

	var fromCharCode = String.fromCharCode;
	// encoder stuff
	var cb_utob = function (c) {
		if (c.length < 2) {
			var cc = c.charCodeAt(0);
			return cc < 0x80 ? c :
				cc < 0x800 ? (fromCharCode(0xc0 | (cc >>> 6)) +
					fromCharCode(0x80 | (cc & 0x3f))) :
				(fromCharCode(0xe0 | ((cc >>> 12) & 0x0f)) +
					fromCharCode(0x80 | ((cc >>> 6) & 0x3f)) +
					fromCharCode(0x80 | (cc & 0x3f)));
		} else {
			var cc = 0x10000 +
				(c.charCodeAt(0) - 0xD800) * 0x400 +
				(c.charCodeAt(1) - 0xDC00);
			return (fromCharCode(0xf0 | ((cc >>> 18) & 0x07)) +
				fromCharCode(0x80 | ((cc >>> 12) & 0x3f)) +
				fromCharCode(0x80 | ((cc >>> 6) & 0x3f)) +
				fromCharCode(0x80 | (cc & 0x3f)));
		}
	};
	var re_utob = /[\uD800-\uDBFF][\uDC00-\uDFFFF]|[^\x00-\x7F]/g;
	var utob = function (u) {
		return u.replace(re_utob, cb_utob);
	};
	var _encode = function (u) {
		var isUint8Array = Object.prototype.toString.call(u) === '[object Uint8Array]';
		if (isUint8Array)
			return u.toString('base64')
		else
			return btoa(utob(String(u)));
	}

	if (typeof global.btoa !== 'function') global.btoa = func_btoa;

	function func_btoa(input) {
		var str = String(input);
		var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
		for (
			// initialize result and counter
			var block, charCode, idx = 0, map = chars, output = '';
			// if the next str index does not exist:
			//   change the mapping table to "="
			//   check if d has no fractional digits
			str.charAt(idx | 0) || (map = '=', idx % 1);
			// "8 - idx % 1 * 8" generates the sequence 2, 4, 6, 8
			output += map.charAt(63 & block >> 8 - idx % 1 * 8)
		) {
			charCode = str.charCodeAt(idx += 3 / 4);
			if (charCode > 0xFF) {
				throw new InvalidCharacterError("'btoa' failed: The string to be encoded contains characters outside of the Latin1 range.");
			}
			block = block << 8 | charCode;
		}
		return output;
	}

	var encode = function (u, urisafe) {
		return !urisafe ?
			_encode(u) :
			_encode(String(u)).replace(/[+\/]/g, function (m0) {
				return m0 == '+' ? '-' : '_';
			}).replace(/=/g, '');
	};

	function IEVersion() {
		var userAgent = navigator.userAgent; //取得浏览器的userAgent字符串  
		var isIE = userAgent.indexOf("compatible") > -1 && userAgent.indexOf("MSIE") > -1; //判断是否IE<11浏览器  
		var isEdge = userAgent.indexOf("Edge") > -1 && !isIE; //判断是否IE的Edge浏览器  
		var isIE11 = userAgent.indexOf('Trident') > -1 && userAgent.indexOf("rv:11.0") > -1;
		if (isIE) {
			var reIE = new RegExp("MSIE (\\d+\\.\\d+);");
			reIE.test(userAgent);
			var fIEVersion = parseFloat(RegExp["$1"]);
			if (fIEVersion == 7) {
				return 7;
			} else if (fIEVersion == 8) {
				return 8;
			} else if (fIEVersion == 9) {
				return 9;
			} else if (fIEVersion == 10) {
				return 10;
			} else {
				return 6; //IE版本<=7
			}
		} else if (isEdge) {
			return 20; //edge
		} else if (isIE11) {
			return 11; //IE11  
		} else {
			return 30; //不是ie浏览器
		}
	}

	function WpsStart(clientType, startInfo, callback) {
		var strData = JSON.stringify(startInfo);
		if (IEVersion() < 10)
			eval("strData = '" + JSON.stringify(startInfo) + "';");
		var baseData = encode(strData);
		startWps(clientType, baseData, callback);
	}

	function WpsStartWrap(clientType, name, func, param, callback) {
		var startInfo = {
			"name": name,
			"function": func,
			"info": param
		};
		WpsStart(clientType, startInfo, callback);
	}

	global.WpsStartUp = {
		StartUp: WpsStartWrap,
		ClientType: {
			wps: "wps",
			et: "et",
			wpp: "wpp"
		}
	}

	return {
		WpsStartUp: global.WpsStartUp
	}
}));