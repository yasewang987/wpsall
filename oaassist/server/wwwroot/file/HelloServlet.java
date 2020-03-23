package com.kso.test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

@WebServlet({ "/HelloServlet" })
public class HelloServlet extends HttpServlet {
    private static final long serialVersionUID = 1L;
    private static String s_fileName = "";

    ByteArrayOutputStream output;
    /**
     * 下载文件
     * @param request
     * @param response
     * @throws ServletException
     * @throws IOException
     */
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        System.out.println("--------doGet--------");
        Cookie[] cookies = request.getCookies();
        String name = "";
        String value = "";
        if (cookies == null || cookies.length == 0) {

            response.sendError(404, "no ssison");
        } else {

            Cookie cookie = cookies[0];
            name = cookie.getName();
            value = cookie.getValue();
            System.out.println(request.getRequestURL());
        }
        System.out.println("JSESSIONID=" + value);
        response.getWriter().write(name + "=" + value);
    }
    /**
     * 上传文件
     * @param request
     * @param response
     * @throws ServletException
     * @throws IOException
     */
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        request.setCharacterEncoding("utf-8");
        response.setContentType("text/html;charset=utf-8");
        System.out.println("添加任务");
        HttpSession session = request.getSession();
        Cookie[] cookies = request.getCookies();
        String value = "";

        try {
            DiskFileItemFactory factory = new DiskFileItemFactory();
            ServletFileUpload upload = new ServletFileUpload((FileItemFactory) factory);
            upload.setHeaderEncoding("UTF-8");

            List items = upload.parseRequest(request);
            boolean isOk = false;
            Map<Object, Object> param = new HashMap<>();
            for (Object object : items) {
                FileItem fileItem = (FileItem) object;
                if (fileItem.isFormField()) {
                    System.out.println(fileItem.getFieldName() + ":" + fileItem.getString("utf-8") + ", size:"
                            + fileItem.getSize());
                    param.put(fileItem.getFieldName(), fileItem.getString("utf-8"));
                    continue;
                }
                String fieldName = fileItem.getFieldName();
                String fileName = fileItem.getName();
                if (fileName.equals("blob"))
                    if (param.containsKey("filename")) {//关键就是要在前端传来的数据中有这个字段
                        fileName = param.get("filename").toString();
                    } else if (param.containsKey("fileName")) {
                        fileName = param.get("fileName").toString();
                    }
                String filePath = request.getSession().getServletContext().getRealPath("/") + fileName;
                System.out.println(fieldName + ":" + filePath);

                FileOutputStream fileOut = new FileOutputStream(filePath);
                InputStream in = fileItem.getInputStream();
                byte[] buffer = new byte[1024];
                int len = 0;
                while ((len = in.read(buffer)) > 0) {
                    fileOut.write(buffer, 0, len);
                }
                in.close();
                fileOut.close();
                response.setHeader("Content-disposition", "attachment; filename*=UTF-8''" + fileName);
                response.getWriter().write(fileName.concat("上传成功"));

                return;
            }
        } catch (FileUploadException e) {

            e.printStackTrace();
        }

        response.sendError(404, "no ssison");
    }
}