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

import com.wps.api.tree.et.*;
import com.wps.runtime.utils.WpsArgs;
import com4j.Variant;
import sun.awt.WindowIDProvider;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Project              :   Java API Examples
 * Comments             :   WPS表格标签页面，包含一个左侧标签栏（LeftMenuPanel）和右侧的wps显示区域（OfficePanel）
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */

public class EtMainPanel extends JPanel {

    private LeftMenuPanel menuPanel;
    private OfficePanel officePanel;
    private Application app = null;
    private static final int DEFAULT_LCID = 2052; //默认的lcid,遇到有lcid的接口传这个就可以.

    public EtMainPanel() {
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
                WpsArgs args = WpsArgs.ARGS_MAP.get(WpsArgs.App.ET);
                args.setWinid(pid.getWindow());
                args.setHeight(client.getHeight());
                args.setWidth(client.getWidth());
//                args.setCrypted(false); //wps2016需要关闭加密
                app = ClassFactory.createApplication();
            }
        });


        menuPanel.addButton("常用", "创建新表格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_Workbooks().Add(Variant.getMissing(), DEFAULT_LCID);
            }
        });

        menuPanel.addButton("常用", "关闭", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                app.get_ActiveWorkbook().Close(false, Variant.getMissing(), Variant.getMissing(), DEFAULT_LCID);
            }
        });

        menuPanel.addButton("常用", "获取版本", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String version = String.valueOf(app.get_Build(DEFAULT_LCID));
                JOptionPane.showMessageDialog(null, version);
            }
        });

        menuPanel.addButton("常用", "字体加粗", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Workbook workbook = app.get_ActiveWorkbook();
                Worksheet sheet = (Worksheet)workbook.get_ActiveSheet();
                sheet.get_Range("A1", "B5").get_Font().put_Bold(true);
            }
        });

        menuPanel.addButton("常用", "设置字号", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Workbook workbook = app.get_ActiveWorkbook();
                Worksheet sheet = (Worksheet)workbook.get_ActiveSheet();
                sheet.get_Range("A1", "A1").get_Font().put_Size(20);
            }
        });

        menuPanel.addButton("常用", "单元格边框", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Workbook workbook = app.get_ActiveWorkbook();
                Worksheet sheet = (Worksheet)workbook.get_ActiveSheet();
                sheet.get_Range("B2", "C9")._BorderAround(5, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, Variant.getMissing());
            }
        });

        menuPanel.addButton("常用", "单元格水平居中", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                Workbook workbook = app.get_ActiveWorkbook();
                Worksheet sheet = (Worksheet)workbook.get_ActiveSheet();
                sheet.get_Range("B2", "C9").put_HorizontalAlignment(-4108);
            }
        });

        menuPanel.addButton("常用", "合并单元格", new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
               Worksheet sheet = (Worksheet)app.get_ActiveSheet();
               sheet.get_Range("A1", "B4").Merge(Variant.getMissing());
            }
        });
    }
}
