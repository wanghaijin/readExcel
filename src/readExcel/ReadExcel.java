package readExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	public Object readExcel(String path) throws IOException {
		if(path == null || Common.EMPTY.equals(path)) {
			return null;
		}else {
			String postfix=Util.getPostfix(path);
			if(!Common.EMPTY.equals(postfix)) {
				if(Common.OFFICE_2003.equals(postfix)) {
					return readXls(path);
				}else if (Common.OFFICE_2010.equals(postfix)) {
					return readXlsx(path);
				}else {
					System.out.println(path+" "+ Common.NOT_EXCEL_FILE);
				}
			}
			return null;
		}
		
	}
	

	// 获取Excel 97-2007
	private static Object readXls(String path) throws IOException {
		InputStream iStream = new FileInputStream(path);
		//Excel文件
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(iStream);
		
		//存储Excel
		List<Object[][]> list=new ArrayList<>();
		//读取sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			//存储sheet数据
			List<List<Object>> lists=new ArrayList<List<Object>>();
			if (hssfSheet == null) {
				continue;
			}
			//行数,此函数返回的是去除第一行后的行数,加一后为全部函数
			int maxRowNum=hssfSheet.getLastRowNum()+1;
			//列数
			int maxCellNum=hssfSheet.getRow(0).getLastCellNum();
			//存储一个表数据
			String[][] strings=new String[maxRowNum][maxCellNum];
			for (int rowNum = 0; rowNum < maxRowNum; rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);

				if (hssfRow != null) {
					//读取每列的数据
					for(int cellNum=0;cellNum<maxCellNum;cellNum++) {
						String string=Util.getValue(hssfRow.getCell(cellNum));
						strings[rowNum][cellNum]=string;
					}
				}
			}
			list.add(strings);
		}
		return list;
	}
	
	//获取Excel2010-数据
	private static Object readXlsx(String path) throws IOException {
		InputStream iStream=new FileInputStream(path);
		XSSFWorkbook xssfWorkbook=new XSSFWorkbook(iStream);
		List<Object[][]> list=new ArrayList<>();
		int maxSheetNum=xssfWorkbook.getNumberOfSheets();
		
		for(int numSheet=0;numSheet<maxSheetNum;numSheet++) {
			XSSFSheet xssfSheet=xssfWorkbook.getSheetAt(numSheet);
			if (xssfSheet==null) {
				continue;
			}
			//行数,此函数返回的是去除第一行后的行数,加一后为全部函数
			int maxRowNum=xssfSheet.getLastRowNum()+1;
			//列数
			int maxCellNum=xssfSheet.getRow(0).getLastCellNum();
			String [][] strings=new String[maxRowNum][maxCellNum];;
			for(int rowNum=0;rowNum<maxRowNum;rowNum++) {
				XSSFRow xssfRow=xssfSheet.getRow(rowNum);
				//获取行数
				maxCellNum=xssfRow.getLastCellNum();
				if (xssfRow!=null) {
					//读取行数据,并将数据导入list
					for(int cellNum=0;cellNum<maxCellNum;cellNum++) {
						String string=Util.getValue(xssfRow.getCell(cellNum));
						strings[rowNum][cellNum]=string;
					}
					
				}	
			}
			list.add(strings);
		}
		return list;
	}

}
