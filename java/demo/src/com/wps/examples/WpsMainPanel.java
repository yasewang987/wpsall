/*
** Copyright @ 2012-2019, Kingsoft office,All rights reserved.
** Redistribution and use in source and binary forms, with or without 
** modification, are permitted provided that the following conditions are met:
**
** 1. Redistributions of source code must retain the above copyright notice, 
**    this list of conditions and the following disclaimer.
** 2. Redistributions in binary form must reproduce the above copyright notice,
**    this list of conditions and the following disclaimer in the documentation
**    and/or other materials provided with the distribution.
** 3. Neither the name of the copyright holder nor the names of its 
**    contributors may be used to endorse or promote products derived from this
**    software without specific prior written permission.
**
** THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
** AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
** IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
** ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE 
** LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
** CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
** SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
** INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
** CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
** ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
** POSSIBILITY OF SUCH DAMAGE.
*/
package com.wps.examples;

import com.wps.api.tree.kso.*;
import com.wps.api.tree.kso.events._CommandBarButtonEvents;
import com.wps.api.tree.wps.*;
import com.wps.api.tree.wps.ClassFactory;
import com.wps.api.tree.wps.events.ApplicationEvents4;
import com.wps.api.tree.ex._ApplicationEx;
import com.wps.api.tree.ex._DocumentEx;
import com.wps.runtime.utils.WpsArgs;
import com4j.EventCookie;
import com4j.Holder;
import com4j.Variant;

import org.apache.commons.codec.binary.Base64;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import sun.awt.WindowIDProvider;
import sun.awt.X11.XEmbedCanvasPeer;

import javax.swing.*;
import java.awt.*;
import java.awt.FileDialog;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;


/**
 * Project              :   Java API Examples
 * Comments             :   WPS文字标签页面，包含一个左侧标签栏（LeftMenuPanel）和右侧的wps显示区域（OfficePanel）
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class WpsMainPanel extends JPanel {

    private LeftMenuPanel menuPanel;
    private OfficePanel officePanel;
    private Application app = null;
    private EventCookie cookie = null;
    private EventCookie cookie2 = null;
    private static Logger logger = LoggerFactory.getLogger(WpsMainPanel.class);
    private CommandBarControl control = null;

    public WpsMainPanel() {
        this.setLayout(new BorderLayout());
        menuPanel = new LeftMenuPanel();
        officePanel = new OfficePanel();
        this.add(menuPanel, BorderLayout.WEST);
        this.add(officePanel, BorderLayout.CENTER);
        initNormalMenu();           //常用接口
        initLocalMenu();            //本地文档相关接口
        initMenu();                 //菜单栏操作接口
        initRepairedMenu();         //修订接口
        initEventMenu();            //事件监听回调接口
        initOthersMenu();           //其他
    }

    public static String getPath(String title, int type){
        FileDialog dialog = new FileDialog((JFrame)null, title, type);
        dialog.setVisible(true);
        if(dialog.getDirectory() == null || dialog.getFile() == null)
            throw new RuntimeException("选择的文件不能为空!");
        return dialog.getDirectory() + dialog.getFile();
    }

    private void initNormalMenu(){
        menuPanel.addButton("常用", "初始化", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Canvas client = officePanel.getCanvas();
                if (app != null) {
                    if(null != app.get_ActiveWindow()){
                        JOptionPane.showMessageDialog(client, "已经初始化过，不需要重新初始化！");
                        return;
                    }
                }
                WindowIDProvider pid = (WindowIDProvider) client.getPeer();
                XEmbedCanvasPeer peer = (XEmbedCanvasPeer) pid;
                peer.removeXEmbedDropTarget();
                try {
                    Method detachChiled = XEmbedCanvasPeer.class.getDeclaredMethod("detachChild");
                    detachChiled.setAccessible(true);
                    Method isXEmbedActive = XEmbedCanvasPeer.class.getDeclaredMethod("isXEmbedActive");
                    isXEmbedActive.setAccessible(true);
                    Boolean isActive = (Boolean) isXEmbedActive.invoke(peer);
                    if(isActive){
                        detachChiled.invoke(peer);
                    }
                } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException ex) {
                    ex.printStackTrace();
                }
                WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPS);
//              args.setPath("/home/wps/workspace/wps_2016/build_debug/WPSOffice/office6/wps"); //手动指定的wps程序路径（默认调用/usr/bin/wps）
                args.setWinid(pid.getWindow());
                args.setHeight(client.getHeight());
                args.setWidth(client.getWidth());
//                args.setCrypted(false); //wps2016需要关闭加密
                app = ClassFactory.createApplication();
            }
        });


        menuPanel.addButton("常用", "创建新文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Add(Variant.getMissing(), Variant.getMissing(), Variant.getMissing(), Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "获取版本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String version = app.get_Build();
                JOptionPane.showMessageDialog(null, version);
            }
        });

        menuPanel.addButton("常用", "关闭", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Close(false, Variant.getMissing(), Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "打印文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "在当前光标处插入图片", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_InlineShapes().AddPicture(getPath("请选择一张图片", FileDialog.LOAD),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "当前光标处插入图片_浮于文字上方", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Document pDoc = app.get_ActiveDocument();
                String imagePath =getPath("请选择一张图片", FileDialog.LOAD);
                ImageIcon image = new ImageIcon(imagePath);
                double width = image.getIconWidth();
                double height = image.getIconHeight();
                double left = Double.parseDouble(app.get_Selection().get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary).toString());
                double top = Double.parseDouble(app.get_Selection().get_Information(WdInformation.wdVerticalPositionRelativeToTextBoundary).toString());
                int page = Integer.parseInt(app.get_Selection().get_Information(WdInformation.wdActiveEndPageNumber).toString());

                Object range = pDoc.GoTo(WdGoToItem.wdGoToPage,WdGoToDirection.wdGoToAbsolute,page,Variant.getMissing());
                Variant anchor = new Variant();
                anchor.set(range);
                double pageTop = 0.0;
                View pView = app.get_ActiveWindow().get_View();
                pView.get_DisplayPageBoundaries();
                if(pView != null && !pView.get_DisplayPageBoundaries()){
                    PageSetup pPageSetup = pDoc.get_PageSetup();
                    pageTop = pPageSetup.get_TopMargin();
                }
                pDoc.get_Shapes().AddPicture(
                        imagePath,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        left,
                        top - pageTop,
                        width,
                        height,
                        anchor
                );
            }
        });

        menuPanel.addButton("常用", "插入图片_带尺寸坐标", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Document pDoc = app.get_ActiveDocument();
                double pageTop = 0.0;
                View pView = app.get_ActiveWindow().get_View();
                pView.get_DisplayPageBoundaries();
                if(pView != null && !pView.get_DisplayPageBoundaries()){
                    PageSetup pPageSetup = pDoc.get_PageSetup();
                    pageTop = pPageSetup.get_TopMargin();
                }
                pDoc.get_Shapes().AddPicture(getPath("请选择一张图片", FileDialog.LOAD), //10, 20, 80, 100
                        Variant.getMissing(),
                        Variant.getMissing(),
                        (double)10,
                        (double)20 - pageTop,
                        80,
                        100,
                        Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "获取文本内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String text = app.get_ActiveDocument().get_Content().get_Text();
                text.replace("/[\r]/g","\r\n");
                JOptionPane.showMessageDialog(null, "<html><body><p style='width: 200px;'>"+text+"</p></body></html>");//文本内容折行显示
            }
        });

        menuPanel.addButton("常用", "设置审批用户名", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.put_UserName("shenpiyonghu123");
            }
        });

        menuPanel.addButton("常用", "获取审批用户名", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JOptionPane.showMessageDialog(null, "<html><body><p style='width: 200px;'>"+app.get_UserName()+"</p></body></html>");
            }
        });

        menuPanel.addButton("常用", "设置标题", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Selection selection = app.get_Selection();
                selection.put_Style("标题 1");
                selection.TypeText("设置标题");
            }
        });

        menuPanel.addButton("常用", "设置字体", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Selection selection = app.get_Selection();
                selection.get_Font().put_Name("Dotum");
                selection.TypeText("设置字体为Dotum");
            }
        });

        menuPanel.addButton("常用", "设置字体大小", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Selection selection = app.get_Selection();
                selection.get_Font().put_Size(36);
                selection.TypeText("设置字体大小为36");
            }
        });

        menuPanel.addButton("常用", "添加编号", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Selection selection = app.get_Selection();
                selection.TypeText("编号一");
                ListTemplate listT = app.get_ListGalleries().Item(WdListGalleryType.wdNumberGallery).get_ListTemplates().Item(1);
                selection.get_Range().get_ListFormat().ApplyListTemplate(listT, WdContinue.wdContinueList, WdListApplyTo.wdListApplyToSelection, 262144);
            }
        });

        menuPanel.addButton("常用", "插入分页符", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().InsertBreak(WdBreakType.wdPageBreak);
            }
        });

        menuPanel.addButton("常用", "插入超链接", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String path = getPath("请选择一个文字文件", FileDialog.LOAD);
                app.get_ActiveDocument().get_Hyperlinks().Add(app.get_Selection().get_Range(), path, "", "", path, "");
            }
        });

        menuPanel.addButton("常用", "插入目录", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                for (int i = 1; i < 4; i++)
                {
                    String title = i + "级标题";
                    String lever = "标题 " + i;
                    app.get_Selection().TypeText(title);
                    app.get_Selection().put_Style(lever);
                    app.get_Selection().TypeParagraph();
                }

                app.get_Selection().SetRange(0,0);
                Range rang = app.get_Selection().get_Range();

                boolean UseHeadingStyles = true;
                long UpperHeadingLevel = 1;
                long LowerHeadingLevel = 3;
                app.get_ActiveDocument().get_TablesOfContents().Add(rang, true, UpperHeadingLevel, LowerHeadingLevel,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                        );
            }
        });

        menuPanel.addButton("常用", "插入页眉页脚", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View view = app.get_ActiveWindow().get_ActivePane().get_View();
                view.put_SeekView(WdSeekView.wdSeekCurrentPageHeader);
                app.get_Selection().TypeText("插入页眉");

                view.put_SeekView(WdSeekView.wdSeekCurrentPageFooter);
                app.get_Selection().TypeText("插入页脚");
            }
        });
        menuPanel.addButton("常用", "打开文档结构图", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveWindow().put_DocumentMap(true);
            }
        });

        menuPanel.addButton("常用", "分栏", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String text = "君不见黄河之水天上来，奔流到海不复回。" +
                        "君不见高堂明镜悲白发，朝如青丝暮成雪。人生得意须尽欢，莫使金樽空对月。" +
                        "天生我材必有用，千金散尽还复来。烹羊宰牛且为乐，会须一饮三百杯。" +
                        "岑夫子，丹丘生，将进酒，杯莫停。与君歌一曲，请君为我倾耳听。" +
                        "钟鼓馔玉不足贵，但愿长醉不愿醒。古来圣贤皆寂寞，惟有饮者留其名。" +
                        "陈王昔时宴平乐，斗酒十千恣欢谑。主人何为言少钱，径须沽取对君酌。" +
                        "五花马、千金裘，呼儿将出换美酒，与尔同销万古愁。";
                int textLen = text.length();
                Selection selection = app.get_Selection();
                selection.TypeText(text);
                selection.InsertBreak(3);

                selection.SetRange(0, textLen);
                TextColumns textColumns = selection.get_Range().get_PageSetup().get_TextColumns();
                textColumns.SetCount(3);
                textColumns.put_EvenlySpaced(0);
            }
        });

        menuPanel.addButton("常用", "清除格式", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().ClearFormatting();
            }
        });

        menuPanel.addButton("常用", "旋转图形", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                com.wps.api.tree.wps.Shapes shapes = app.get_ActiveDocument().get_Shapes();
                int Type = (int)MsoAutoShapeType.msoShapeCan.comEnumValue();
                float Left = 150;
                float Top = 122;
                float Width = 118;
                float Height = 83;
                com.wps.api.tree.wps.Shape shape = shapes.AddShape(Type, Left, Top, Width, Height, Variant.getMissing());
                shape.IncrementRotation(-90);
            }
        });


    }

    private  void initLocalMenu(){
        menuPanel.addButton("本地文档", "打开可编辑本地文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD),
                        Variant.getMissing(),
                        false,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "打开可编辑本地文档_通过文档对话框", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc=new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.showDialog(new JLabel(), "打开");
                String fileName = jfc.getSelectedFile().getPath();

                app.get_Documents().Open(fileName,
                        Variant.getMissing(),
                        false,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "只读模式打开本地文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD),
                        Variant.getMissing(),
                        true,
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                ).Protect(WdProtectionType.wdAllowOnlyReading,Variant.getMissing(),
                        Variant.getMissing(),Variant.getMissing(),Variant.getMissing());
            }
        });


        menuPanel.addButton("本地文档", "保存本地弹窗", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                HashMap formatMap = new HashMap<String,Integer>();
                formatMap.put("doc",0);
                formatMap.put("wps",0);
                formatMap.put("wpt",1);
                formatMap.put("dot",1);
                formatMap.put("txt",7);
                formatMap.put("rtf",6);
                formatMap.put("html",8);
                formatMap.put("htm",8);
                formatMap.put("mhtml",9);
                formatMap.put("xml",11);
                formatMap.put("docx",12);
                formatMap.put("docm",13);
                formatMap.put("dotx",14);
                formatMap.put("dotm",15);
                formatMap.put("uof",100);
                formatMap.put("uot",101);
                formatMap.put("pdf",17);

                //需要打开本地文件对话框
                JFileChooser jfc=new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.showDialog(new JLabel(), "另存为");
                String fileName = jfc.getSelectedFile().getPath();
                String[] formats = fileName.split("\\.");
                String format;
                if(formats.length == 0 || !formatMap.containsKey(formats[formats.length - 1])) { //路径不存在.或者后缀名无法识别，默认保存为.wps
                    format = "wps";
                    fileName += ".wps";
                }
                else {
                    format = formats[formats.length - 1];
                }
                // 获取到文件名保存
                Document pDoc = app.get_ActiveDocument();
                pDoc.SaveAs(fileName,
                        formatMap.get(format),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("本地文档", "打开隐藏文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Documents().Open(getPath("请选择一个文字文件", FileDialog.LOAD), false, false, false, "", "", false, "", "", 0, 0, false, false, false, 0, false);
            }
        });

    }

    private void initMenu(){
        menuPanel.addButton("菜单栏", "隐藏文件菜单", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                CommandBar commandBar = app.get_CommandBars().get_Item("Menu Bar");
                if(commandBar == null){
                    logger.error("CommandBar is null!");
                    return ;
                }
                CommandBarControl control = commandBar.get_Controls().get_Item(1);
                if(control == null){
                    logger.error("CommandBarControl is null!");
                    return ;
                }
                control.put_Visible(false);
            }
        });

        menuPanel.addButton("菜单栏", "禁用/启用-接受修订按钮", new ActionListener() {
            private boolean enable = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                enable = !enable;
                app.get_CommandBars().get_Item("Track Changes").get_Controls().get_Item(7).put_Enabled(enable);
                app.get_CommandBars().get_Item("Reviewing").get_Controls().get_Item(5).put_Enabled(enable);
            }
        });

        menuPanel.addButton("菜单栏", "禁用/启用-拒绝修订按钮", new ActionListener() {
            private boolean enable = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                enable = !enable;
                app.get_CommandBars().get_Item("Track Changes").get_Controls().get_Item(8).put_Enabled(enable);
                app.get_CommandBars().get_Item("Reviewing").get_Controls().get_Item(6).put_Enabled(enable);

            }
        });

        menuPanel.addButton("菜单栏", "显示工具条", new ActionListener() {
            private HashMap<String,Integer> m_BarCatchMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                m_BarCatchMap.put("Standard", 0);
                m_BarCatchMap.put("Formatting", 0);
                m_BarCatchMap.put("Tables and Borders", 0);
                m_BarCatchMap.put("Align", 0);
                m_BarCatchMap.put("Reviewing", 0);
                m_BarCatchMap.put("Extended Formatting", 0);
                m_BarCatchMap.put("Drawing", 0);
                m_BarCatchMap.put("Symbol", 0);
                m_BarCatchMap.put("3-D Settings", 0);
                m_BarCatchMap.put("Shadow Settings", 0);
                m_BarCatchMap.put("Mail Merge", 0);
                m_BarCatchMap.put("Control Toolbox", 0);
                m_BarCatchMap.put("Outlining", 0);
                m_BarCatchMap.put("Forms", 0);
                app.get_CommandBars().get_Item("menu bar").put_Enabled(true);

                for (String key : m_BarCatchMap.keySet()) {
                    app.get_CommandBars().get_Item(key).put_Enabled(true);
                }
            }
        });

        menuPanel.addButton("菜单栏", "隐藏工具条", new ActionListener() {
            private HashMap<String,Integer> m_BarCatchMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                m_BarCatchMap.put("Standard", 0);
                m_BarCatchMap.put("Formatting", 0);
                m_BarCatchMap.put("Tables and Borders", 0);
                m_BarCatchMap.put("Align", 0);
                m_BarCatchMap.put("Reviewing", 0);
                m_BarCatchMap.put("Extended Formatting", 0);
                m_BarCatchMap.put("Drawing", 0);
                m_BarCatchMap.put("Symbol", 0);
                m_BarCatchMap.put("3-D Settings", 0);
                m_BarCatchMap.put("Shadow Settings", 0);
                m_BarCatchMap.put("Mail Merge", 0);
                m_BarCatchMap.put("Control Toolbox", 0);
                m_BarCatchMap.put("Outlining", 0);
                m_BarCatchMap.put("Forms", 0);
                app.get_CommandBars().get_Item("menu bar").put_Enabled(false);

                for (String key : m_BarCatchMap.keySet()) {
                    app.get_CommandBars().get_Item(key).put_Enabled(false);
                }
            }
        });
    }

    private void initRepairedMenu(){
        menuPanel.addButton("修订", "开启/关闭修订", new ActionListener() {
            private boolean m_status = false;
            @Override
            public void actionPerformed(ActionEvent e) {
                m_status = !m_status;
                app.get_ActiveDocument().put_TrackRevisions(m_status);
            }
        });

        menuPanel.addButton("修订", "显示标记的最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(true);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);

            }
        });

        menuPanel.addButton("修订", "显示原始状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewOriginal);

            }
        });

        menuPanel.addButton("修订", "显示最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);

            }
        });

        menuPanel.addButton("修订", "打印修订", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(true);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "打印原始状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewOriginal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "打印最终状态", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                View pView = app.get_ActiveWindow().get_View();
                if(pView == null){
                    logger.error("View is null");
                    return ;
                }
                pView.put_ShowRevisionsAndComments(false);
                pView.put_RevisionsView(WdRevisionsView.wdRevisionsViewFinal);
                ((Document)app.get_ActiveDocument()).PrintOut(
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing(),
                        Variant.getMissing()
                );
            }
        });

        menuPanel.addButton("修订", "保护文档", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Protect(WdProtectionType.wdAllowOnlyReading,Variant.getMissing(),Variant.getMissing(),Variant.getMissing(),Variant.getMissing());
            }
        });

        menuPanel.addButton("修订", "停止保护", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().Unprotect(Variant.getMissing());
            }
        });

        menuPanel.addButton("修订", "禁止剪切", new ActionListener() {
            private HashMap<String,Integer> m_enableCutMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                //cut
                m_enableCutMap.put("Chart Area Popup Menu", 2);
                m_enableCutMap.put("Dml Shapes", 2);
                m_enableCutMap.put("WordArt Context Menu", 1);
                m_enableCutMap.put("Tables", 1);
                m_enableCutMap.put("Track Changes", 1);
                m_enableCutMap.put("Inline Picture", 1);
                m_enableCutMap.put("Lists", 1);
                m_enableCutMap.put("Frames", 1);
                m_enableCutMap.put("Field Display List Numbers", 1);
                m_enableCutMap.put("Fields", 1);
                m_enableCutMap.put("Shapes", 1);
                m_enableCutMap.put("Text", 1);
                m_enableCutMap.put("Spelling Suggestions Popup Menu", 5);
                m_enableCutMap.put("Edit", 3);
                m_enableCutMap.put("Table Text", 1);
                m_enableCutMap.put("Table Cells", 1);
                m_enableCutMap.put("Whole Table", 1);
                m_enableCutMap.put("Picture Context Menu", 2);
                m_enableCutMap.put("Numbered Popup Menu", 2);
                m_enableCutMap.put("Bulleted Popup Menu", 2);
                m_enableCutMap.put("Standard", 11);

                for (String key : m_enableCutMap.keySet()) {
                    app.get_CommandBars().get_Item(key).get_Controls().get_Item(m_enableCutMap.get(key)).put_Enabled(false);
                }
            }
        });

        menuPanel.addButton("修订", "禁止复制", new ActionListener() {
            private HashMap<String,Integer> m_enableCopyMap = new HashMap<String, Integer>();
            @Override
            public void actionPerformed(ActionEvent e) {
                //copy
                m_enableCopyMap.put("Chart Area Popup Menu", 1);
                m_enableCopyMap.put("Dml Shapes", 1);
                m_enableCopyMap.put("WordArt Context Menu", 2);
                m_enableCopyMap.put("Tables", 2);
                m_enableCopyMap.put("Track Changes", 2);
                m_enableCopyMap.put("Inline Picture", 2);
                m_enableCopyMap.put("Lists", 2);
                m_enableCopyMap.put("Frames", 2);
                m_enableCopyMap.put("Field Display List Numbers", 1);
                m_enableCopyMap.put("Fields", 2);
                m_enableCopyMap.put("Shapes", 2);
                m_enableCopyMap.put("Text", 2);
                m_enableCopyMap.put("Spelling Suggestions Popup Menu", 4);
                m_enableCopyMap.put("Edit", 4);
                m_enableCopyMap.put("Table Text", 2);
                m_enableCopyMap.put("Table Cells", 2);
                m_enableCopyMap.put("Whole Table", 2);
                m_enableCopyMap.put("Picture Context Menu", 1);
                m_enableCopyMap.put("Numbered Popup Menu", 1);
                m_enableCopyMap.put("Bulleted Popup Menu", 1);
                m_enableCopyMap.put("Standard", 12);

                for (String key : m_enableCopyMap.keySet()) {
                    app.get_CommandBars().get_Item(key).get_Controls().get_Item(m_enableCopyMap.get(key)).put_Enabled(false);
                }
            }
        });

    }

    private void initEventMenu(){
        //事件监听及回调
        menuPanel.addButton("事件监听及回调", "注册关闭事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            if (cookie != null)
                return;
            cookie = app.advise(ApplicationEvents4.class, new ApplicationEvents4() {
                @Override
                public void DocumentBeforeClose(Document doc, Holder<Boolean> cancel) {
                    logger.info("调用注册的关闭事件");
                    JOptionPane.showMessageDialog(null,"关闭事件被拦截");
                    cancel.value = true;
                }
            });
            }
        });

        menuPanel.addButton("事件监听及回调", "取消注册关闭事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (cookie == null)
                    return;
                cookie.close();
                cookie = null;
            }
        });

        menuPanel.addButton("事件监听及回调", "添加一个按钮", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control != null){
                    JOptionPane.showMessageDialog(null,"按钮已存在");
                    return;
                }
                CommandBar bar = app.get_CommandBars().Add("test",1, Variant.getMissing(), Variant.getMissing());
                control = bar.get_Controls().Add(1,1,"test",1,"test");
                bar.put_Visible(true);
                control.put_Visible(true);
                control.put_Caption("test");
                JOptionPane.showMessageDialog(null,"按钮添加成功");
            }
        });

        menuPanel.addButton("事件监听及回调", "注册按钮事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control == null){
                    JOptionPane.showMessageDialog(null,"请先添加按钮");
                    return;
                }
                if(cookie2 != null){
                    JOptionPane.showMessageDialog(null,"事件已注册");
                    return;
                }
                cookie2 = control.advise(_CommandBarButtonEvents.class, new _CommandBarButtonEvents() {
                    @Override
                    public void Click(CommandBarButton ctrl, Holder<Boolean> cancelDefault) {
                        JOptionPane.showMessageDialog(null,"你点击了按钮 test");
                        cancelDefault.value = true;
                    }
                });
                JOptionPane.showMessageDialog(null,"注册事件成功");
            }
        });

        menuPanel.addButton("事件监听及回调", "取消注册按钮事件", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if(control == null){
                    JOptionPane.showMessageDialog(null,"请先添加按钮");
                    return;
                }
                cookie2.close();
                cookie2 = null;
                JOptionPane.showMessageDialog(null,"取消注册事件成功");
            }
        });

    }

    private void initOthersMenu(){

        menuPanel.addButton("其他", "设置打印图像对象", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Options().put_PrintDrawingObjects(true);
            }
        });

        menuPanel.addButton("其他", "设置打印隐藏文字", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Options().put_PrintHiddenText(true);
            }
        });

        menuPanel.addButton("其他", "获取文件大小", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File file = new File(getPath("请选择一个文字文件", FileDialog.LOAD));
                if(file == null){
                    logger.error("file not exist");
                    return ;
                }
                JOptionPane.showMessageDialog(null,file.length());
            }
        });

        menuPanel.addButton("其他", "开启自动备份", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Application().queryInterface(_ApplicationEx.class).put_ForceBackupEnable(true);
            }
        });

        menuPanel.addButton("其他", "关闭自动备份", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Application().queryInterface(_ApplicationEx.class).put_ForceBackupEnable(false);
            }
        });

        menuPanel.addButton("其他", "设置标题行重复", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Rows().put_HeadingFormat(-1);
            }
        });
        menuPanel.addButton("其他", "取消标题行重复", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Rows().put_HeadingFormat(0);
            }
        });

        menuPanel.addButton("其他", "批量插入书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Add("book1",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book2",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book3",Variant.getMissing());
                app.get_Selection().TypeText("，");
                app.get_ActiveDocument().get_Bookmarks().Add("book4",Variant.getMissing());
                app.get_Selection().TypeText("，");
            }
        });

        menuPanel.addButton("其他", "批量删除书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Item("book1").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book2").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book3").Delete();
                app.get_ActiveDocument().get_Bookmarks().Item("book4").Delete();
            }
        });

        menuPanel.addButton("其他", "获取全部书签", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Bookmarks bookmarks = app.get_ActiveDocument().get_Bookmarks();
                String res = new String();
                res += bookmarks.Item("book1").get_Name()+"，";
                res += bookmarks.Item("book2").get_Name()+"，";
                res += bookmarks.Item("book3").get_Name()+"，";
                res += bookmarks.Item("book4").get_Name()+"，";
                JOptionPane.showMessageDialog(null,"<html><body><p style='width: 200px;'>"+res+"</p></body></html>");
            }
        });

        menuPanel.addButton("其他", "替换书签内容", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveDocument().get_Bookmarks().Item("book1").get_Range().put_Text("1111");
                app.get_ActiveDocument().get_Bookmarks().Item("book2").get_Range().put_Text("2222");
                app.get_ActiveDocument().get_Bookmarks().Item("book3").get_Range().put_Text("3333");
                app.get_ActiveDocument().get_Bookmarks().Item("book4").get_Range().put_Text("4444");
            }
        });

        menuPanel.addButton("其他", "从光标位置删除", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().TypeBackspace();
            }
        });

        menuPanel.addButton("其他", "插入表格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Range range = app.get_ActiveDocument().get_Content();
                app.get_ActiveDocument().get_Tables().Add(range,3,5,Variant.getMissing(),Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "适应文字1_选定区域的单元格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Selection().get_Cells().Item(1).put_FitText(true);
            }
        });

        menuPanel.addButton("其他", "按行拆分", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Row row = table.get_Rows().Item(1);
                row.get_Cells().Split(3,2,Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "按列拆分", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Column column = table.get_Columns().Item(2);
                column.get_Cells().Split(3,3,Variant.getMissing());
            }
        });

        menuPanel.addButton("其他", "拆分单元格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Table table = app.get_ActiveDocument().get_Tables().Item(1);
                Row row = table.get_Rows().Item(2);
                row.get_Cells().Item(1).Split(2,3);
            }
        });

    }
}
