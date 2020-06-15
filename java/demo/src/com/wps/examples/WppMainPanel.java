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

import com.wps.api.tree.wpp.*;
import com.wps.runtime.utils.WpsArgs;
import sun.awt.WindowIDProvider;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Project              :   Java API Examples
 * Comments             :   WPS演示标签页面，包含一个左侧标签栏（LeftMenuPanel）和右侧的wps显示区域（OfficePanel）
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class WppMainPanel extends JPanel {

    private LeftMenuPanel menuPanel;
    private OfficePanel officePanel;
    private Application app = null;

    public WppMainPanel() {
        this.setLayout(new BorderLayout());
        menuPanel = new LeftMenuPanel();
        officePanel = new OfficePanel();
        this.add(menuPanel, BorderLayout.WEST);
        this.add(officePanel, BorderLayout.CENTER);
        initMenu();
    }

    private void initMenu(){
        menuPanel.addButton("常用", "初始化", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
            Canvas client = officePanel.getCanvas();
            if (app != null) {
                JOptionPane.showMessageDialog(client, "已经初始化过，不需要重新初始化！");
                return;
            }
            WindowIDProvider pid = (WindowIDProvider) client.getPeer();
            WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.WPP);
//            args.setPath("/home/wps/workspace/wps_2016/build_debug/WPSOffice/office6/wpp"); //手动指定的wpp程序路径（默认调用/usr/bin/wpp）
            args.setWinid(pid.getWindow());
            args.setHeight(client.getHeight());
            args.setWidth(client.getWidth());
//            args.setCrypted(false); //wps2016需要关闭加密
            app = ClassFactory.createApplication();
            }
        });


        menuPanel.addButton("常用", "新建", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Presentations().Add(MsoTriState.msoTrue);
            }
        });

        menuPanel.addButton("常用", "关闭", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActivePresentation().Close();
            }
        });

        menuPanel.addButton("常用", "获取版本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String version = app.get_Build();
                JOptionPane.showMessageDialog(null, version);
            }
        });

        menuPanel.addButton("常用", "设置幻灯片大小", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActivePresentation().get_PageSetup().put_SlideSize(PpSlideSizeType.ppSlideSizeOnScreen16x10);
            }
        });

        menuPanel.addButton("常用", "插入文本框", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                float Left = 100;
                float Top = 100;
                float Width = 300;
                float Height = 200;
                MsoTextOrientation Orientation = MsoTextOrientation.msoTextOrientationHorizontal;
                com.wps.api.tree.wpp.Shape shape = app.get_ActiveWindow().get_Selection().get_SlideRange().get_Shapes().AddTextbox(Orientation, Left, Top, Width, Height);
                shape.get_TextFrame().get_TextRange().put_Text("插入文本框");
            }
        });

        menuPanel.addButton("常用", "文本框居中对齐", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Selection selection = app.get_ActiveWindow().get_Selection();
                float Left = 100;
                float Top = 100;
                float Width = 300;
                float Height = 200;
                MsoTextOrientation Orientation = MsoTextOrientation.msoTextOrientationHorizontal;
                com.wps.api.tree.wpp.Shape shape = selection.get_SlideRange().get_Shapes().AddTextbox(Orientation, Left, Top, Width, Height);
                TextRange textRange = shape.get_TextFrame().get_TextRange();
                textRange.put_Text("文本框居中对齐");
                textRange.Paragraphs(-1, -1).get_ParagraphFormat().put_Alignment(PpParagraphAlignment.ppAlignCenter);
            }
        });

        menuPanel.addButton("常用", "添加页脚", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                HeaderFooter footer = app.get_ActiveWindow().get_Selection().get_SlideRange().get_HeadersFooters().get_Footer();
                if (footer != null)
                {
                    footer.put_Visible(MsoTriState.msoTrue);
                    footer.put_Text("插入页脚");
                }
            }
        });

        menuPanel.addButton("常用", "插入自选图形", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                float Left = 50;
                float Top = 50;
                float Width = 300;
                float Height = 200;
                app.get_ActiveWindow().get_Selection().get_SlideRange().get_Shapes().AddShape(MsoAutoShapeType.msoShape5pointStar, Left, Top, Width, Height);
            }
        });
        menuPanel.addButton("常用", "新增节", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                SectionProperties SectionProperties = app.get_ActivePresentation().get_SectionProperties();
                int count = SectionProperties.get_Count() + 1;
                SectionProperties.AddSection(count,"新增节" + count);
            }
        });
    }
}
