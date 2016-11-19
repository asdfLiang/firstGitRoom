package com.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelUtil {

	public static void readExcel(File file) throws IOException {
		//创建输入流
		FileInputStream fis = new FileInputStream(file);
		//用POIFSFileSystem来解析文件
		POIFSFileSystem ts = new POIFSFileSystem(fis);
		//将一个Excel读入到workbook中
		HSSFWorkbook workbook = new HSSFWorkbook(ts);
		//读取第一个sheet页（sheet页编号从0开始）
		HSSFSheet sh = workbook.getSheetAt(0);

		//循环需要的参数
		int i = 0,j = 0;
		HSSFRow ro = null;
		HSSFCell cell = null;
		
	
		//改变一个单元个中的值
		sh.getRow(4).getCell(2).setCellValue("20");
		
//		/*以下代码为测试代码，在控制台显示读入的数据，需要可自行开启*/
//		//当读完一个Excel的所有行后，结束循环
//		while (sh.getRow(i) != null) {
//			//获取当前行
//			ro = sh.getRow(i);
//			//当读完一行中的所有列后，结束内层循环
//			while( ro.getCell(j) != null){
//				//获取一格Excel单元格的值
//				cell = ro.getCell(j);
//				//输出每一个单元格的值，ro.getRowNum()和cell.getColumnIndex()分别表示行号和列号
//				System.out.print(ro.getRowNum()+"行"+cell.getColumnIndex()+"列："+cell+"\t\t\t");
//				j++;
//			}
//			i++;
//			//此处注意将j的值置为0，下次循环从下一行的第一格的值开始输出
//			j = 0;
//			//输出一行后换行
//			System.out.println();
//		}
//		//读取完输出提示
//		System.out.println("读取完成");
//		/*测试代码结束*/
		
		//用UUID生成随机数，作为新excel文件的名字
		String name = UUID.randomUUID().toString();
		//创建一个向指定位置写入文件的输出流
		FileOutputStream os = new FileOutputStream("D:\\"+name+".xls");
		//向指定的文件写入excel
		workbook.write(os);
		
		//关闭流
		fis.close();
		os.close();
		workbook.close();
	}

}
