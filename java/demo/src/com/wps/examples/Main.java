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

/**
 * Project      : Wps JavaAPI Demo
 * Author       : 金山办公
 * Date         : 2019.12.9
 * Description  : 此项目为WPS二次开发Java版本Demo程序，二次开发接口的使用方式，均可以参照此项目中的代码样例实现。
 * Note         :
 *          1. Java中不提供函数默认参数，因而调用接口时必须给定义的所有参数都传参，如果某个对应位置的参数需要使用默认值，请传递 Variant.getMissing()
 *          2. 获取对象或数值的的接口可以直接判断返回值是否为null进行容错校验，操作类型的接口，容错校验需捕获Exception异常，以此判断操作是否执行成功。
 *          3. 所有对象的属性在设置和获取时，在java中均需使用加上get_或put_前缀的对应方法，无法通过 对象名.属性直接调用。
 */

public class Main {
    public static void main(String []arg) throws InterruptedException, ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {
        java.lang.System.setProperty("sun.awt.xembedserver", "true");           //Linux下必须加这一句才能调用
        //放开下面这段代码可以更换按钮主题风格
//        for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
//            if ("com.sun.java.swing.plaf.gtk.GTKLookAndFeel".equals(info.getClassName())) {
//                javax.swing.UIManager.setLookAndFeel(info.getClassName());
//                break;
//            }
//        }

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                JFrame mainFrame = new JFrame();
                mainFrame.setTitle("WPS JAVA接口调用演示");                       //设置显示窗口标题
                mainFrame.setSize(1524, 768);                           //设置窗口显示尺寸
                mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);       //置窗口是否可以关闭
                mainFrame.setResizable(false);//禁止缩放
                JTabbedPane tabbedPane = new JTabbedPane();
                tabbedPane.add(new WpsMainPanel(), "WPS文字");
                tabbedPane.add(new EtMainPanel(), "WPS表格");
                tabbedPane.add(new WppMainPanel(), "WPS演示");
                mainFrame.add(tabbedPane);
                mainFrame.setVisible(true);                                     //设置窗口是否可见
            }
        });
    }

}
