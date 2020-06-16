# WPS二次开发

这是WPS二次开发的开源代码仓库，采用BSD开源协议，详见[LICENSE](LICENSE)。

## 如果您是第一次刷到这个仓库

[邀你填写「WPS端二次开发支持基本信息调查」](https://f.wps.cn/form-write/4aTnaSY3/)
完成填写后，可以加`3253920855`这个QQ，明确告知需求。

## 如果您想了解现在WPS的最佳实践

请参看[WPS产品矩阵集成解决方案-WPS二次开发集成篇V1.1.pptx](https://kdocs.cn/l/cyKkDebda)

## WPS二次开发场景与本仓库代码对应关系

- WPS加载项集成模式可以实现业务系统与WPS之间「**集成松耦合、业务强关联、功能易扩展、场景全覆盖、切换零成本**」，**回归初心**，让我们互相为对方**赋能**。

- WPS2019 **所有版本** 都支持此种集成方式，包括Windows个人、企业、专业、专业增强，Linux的社区（x86）和企业（各类国产CPU架构）。

- WPS加载项的示例——OA辅助，基本覆盖了业务系统集成WPS的常用场景，**这就是个Demo**，我们希望业务厂商基于此结合自己的使用场景，产出自己的XX助手，将WPS变成自己业务必须的「**超级编辑器**」，共同服务好用户。

- [x] [WPS加载项（jsapi）](oaassist/README.md)【已入库】

---

**以下集成方式需要联系我们的项目经理获取专业版**
- [ ] 嵌入浏览器——ActiveX，IE，Windows环境
- [x] [嵌入浏览器——NPAPI，Firefox/360/奇安信，Windows环境](np-example/browser-integration-wps/README.md)【已入库】
- [x] [嵌入浏览器——NPAPI，Firefox/360/奇安信，Linux环境](np-example/browser-integration-wps/README.md)【已入库】
- [x] [嵌入Java客户端——Windows环境](https://kdocs.cn/l/sJyOAdMPn)【待入库】
- [x] [嵌入Java客户端——Linux环境](java/README.md)【已入库】
- [x] [嵌入C++客户端——Windows环境](https://kdocs.cn/l/c7jl7x76T)【待入库】
- [x] [嵌入C++客户端——Linux环境，QT](cpp/README.md)【已入库】
- [x] [嵌入WPF客户端——Windwos环境，C#](https://kdocs.cn/l/ce4rXtmFS)【待入库】
- [x] [开发Com加载项——Windwos环境，VB/C++](https://kdocs.cn/l/c7jl7x76T)【待入库】
- [x] [开发VSTO加载项——Windows环境，C#](https://kdocs.cn/l/ce4rXtmFS)【待入库】
- [x] [开发加载项——Linxu环境，C++](https://kdocs.cn/l/c1cSaydPa)【待入库】

---

**[WPS开放平台](https://open.wps.cn)上注册成为服务开发商，可以免费对接云文档、在线预览和编辑服务**
- [x] [WebOffice集成（在线预览、编辑、协同）——WPS开发平台，公网](https://open.wps.cn/docs/wwo/join/platform-overview)
  - [各种Demo](https://open.wps.cn/docs/wwo/access/sdk-demo)
  - [API手册](https://wwo.wps.cn/docs-js-sdk/#/)

---

- **联系我们的项目经理获取服务**
- [ ] WebOffice集成（在线预览、编辑、协同）——私有化部署

---

- **只有WPS移动专业版支持二次开发和集成**
- [x] [WPS Office移动专业版集成（AIDL）](http://mo.wps.cn/pc-app/office-pro.html)
- [ ] WPS Office移动预览SDK（仅安卓）

## 怎么用这个仓库

`git clone https://code.aliyun.com/zouyingfeng/wps.git`

设置成了无需注册完全公开可复制的权限

## 想给我们建议？

请通过[issues](https://code.aliyun.com/zouyingfeng/wps/issues)页面向我们提建议，我们会认真查看，并通过此仓库更新发布。

## 想获取WPS的开发资料？

- 所有资料统一发布平台————[WPS开放平台](https://open.wps.cn)
- WPS客户端API资料————[WPS客户端二次开发](https://open.wps.cn/docs/office)
