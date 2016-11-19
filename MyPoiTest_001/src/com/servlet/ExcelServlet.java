package com.servlet;

import java.io.File;
import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.util.ExcelUtil;

public class ExcelServlet extends HttpServlet {

	public ExcelServlet() {
		super();
	}

	public void destroy() {
		super.destroy(); 

	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

	}

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		//获取文件路径
		String path = request.getParameter("excelText");
		//创建文件
		File file = new File(path);
		//调用自定义的工具类完成整个功能
		ExcelUtil.readExcel(file);
	}

	public void init() throws ServletException {
		// Put your code here
	}

}
