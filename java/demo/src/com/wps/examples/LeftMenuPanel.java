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

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.util.HashMap;
import java.util.Map;

/**
 * Project              :   Java API Examples
 * Comments             :   界面左侧第一列标签栏，栏中每一个标签按钮对应一类接口，切换标签时，切换第二列标签栏中对应的接口按钮。维护一个Map，创建时标签数量和内容从Map中动态获取
 * JDK version used     :   JDK1.8
 * Author               :   Kingsoft
 * Create Date          :   2019-12-09
 * Version              :   1.0
 */
public class LeftMenuPanel extends JPanel {

    private JTabbedPane tabbedPane;
    private Map<String, JPanel> panels = new HashMap<>();
    public LeftMenuPanel() {
        this.setLayout(new BorderLayout());
        tabbedPane = new JTabbedPane();
        tabbedPane.setPreferredSize(new Dimension(350, 300));
        tabbedPane.setTabPlacement(JTabbedPane.LEFT);
        this.add(tabbedPane);
    }

    public void addButton(String category, String name, ActionListener listener, boolean enable){
        JPanel panel = null;
        if(panels.containsKey(category)){
            panel = panels.get(category);
        }else{
            panel = new JPanel();
            panel.setLayout(new GridLayout(30,1, 0, 0));
            tabbedPane.add(category, panel);
            panels.put(category, panel);
        }
        JButton button = new JButton(name);
        button.setEnabled(enable);
        button.addActionListener(listener);
        panel.add(button);
    }

    public void addButton(String category, String name, ActionListener listener){
        this.addButton(category,name,listener,true);
    }

}
