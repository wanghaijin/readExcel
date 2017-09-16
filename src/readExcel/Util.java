package readExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Util {
	// 获取文件后缀名
	public static String getPostfix(String path) {
		if (path == null || Common.EMPTY.equals(path.trim())) {
			return Common.EMPTY;
		}
		if (path.contains(Common.POINT)) {
			return path.substring(path.lastIndexOf(Common.POINT) + 1, path.length());
		}

		return Common.EMPTY;
	}

	// 获取Excel 97-2007
	public static List<List<List<Object>>> readXls(String path) throws IOException {
		InputStream iStream = new FileInputStream(path);
		//Excel文件
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(iStream);
		
		//存储Excel
		List<List<List<Object>>> alllist = new ArrayList<List<List<Object>>>();
		//读取sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			//存储sheet数据
			List<List<Object>> lists=new ArrayList<List<Object>>();
			if (hssfSheet == null) {
				continue;
			}
			//行数
			int maxRowNum=hssfSheet.getLastRowNum();
			//列数
			int maxCellNum;
			for (int rowNum = 0; rowNum <= maxRowNum; rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
//				if(rowNum==0) {
					maxCellNum=hssfRow.getLastCellNum();
//				}
				if (hssfRow != null) {
//					model = new Model();
					//存储行数据
					List<Object> list=new ArrayList<Object>();
					//读取每列的数据
					for(int i=0;i<maxCellNum;i++) {
						Object string=getValue(hssfRow.getCell(i));
						list.add(string);
					}
					lists.add(list);
//					HSSFCell no = hssfRow.getCell(0);
//					HSSFCell name = hssfRow.getCell(1);
//					HSSFCell age = hssfRow.getCell(2);
//					HSSFCell score = hssfRow.getCell(4);
//					model.setNo(getValue(no));
//					model.setName(getValue(name));
//					model.setAge(getValue(age));
//					model.setScore(Float.valueOf(getValue(score)));
//					lists.add(model);
				}
			}
			alllist.add(lists);
		}
		return alllist;
	}
	
	//获取Excel2010-数据
	public static List<List<List<Object>>> readXlsx(String path) throws IOException {
		InputStream iStream=new FileInputStream(path);
		XSSFWorkbook xssfWorkbook=new XSSFWorkbook(iStream);
//		Model model=null;
		List<List<List<Object>>> alllist=new ArrayList<List<List<Object>>>();
		int maxSheetNum=xssfWorkbook.getNumberOfSheets();
		for(int numSheet=0;numSheet<maxSheetNum;numSheet++) {
			List<List<Object>> lists=new ArrayList<List<Object>>();
			XSSFSheet xssfSheet=xssfWorkbook.getSheetAt(numSheet);
			if (xssfSheet==null) {
				continue;
			}
			int maxRowNum=xssfSheet.getLastRowNum();
			int maxCellNum;
//			System.out.println(xssfSheet.getLastRowNum());
			for(int rowNum=0;rowNum<=maxRowNum;rowNum++) {
				XSSFRow xssfRow=xssfSheet.getRow(rowNum);
				maxCellNum=xssfRow.getLastCellNum();
				if (xssfRow!=null) {
					List<Object> list=new ArrayList<Object>();
					//读取行数据,并将数据导入list
					for(int i=0;i<maxCellNum;i++) {
						Object string=getValue(xssfRow.getCell(i));
						list.add(string);
					}
					lists.add(list);
				}
//					model=new Model();
//					XSSFCell no=xssfRow.getCell(0);
//					XSSFCell name=xssfRow.getCell(1);
//					XSSFCell age=xssfRow.getCell(2);
//					XSSFCell score=xssfRow.getCell(4);
//					model.setNo(getValue(no));
//					model.setName(getValue(name));
//					model.setAge(getValue(age));
//					model.setScore(Float.valueOf(getValue(score)));
//					lists.add(model);
//				}
			}
			alllist.add(lists);
		}
		return alllist;
	}

	// 获取Excel 97-2007当前行对应列的值
	private static Object getValue(HSSFCell hssfCell) {
		if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
//			return String.valueOf(hssfCell.getBooleanCellValue());
			return hssfCell.getBooleanCellValue();
		}else if(hssfCell.getCellType()==hssfCell.CELL_TYPE_NUMERIC){
//			return String.valueOf(hssfCell.getNumericCellValue());
			return hssfCell.getNumericCellValue();
		}else {
//			return String.valueOf(hssfCell.getRichStringCellValue());
			return hssfCell.getRichStringCellValue();
		}
	}
	
	//获取Excel 2010-当前行对应列的值
	private static Object getValue(XSSFCell xssfRow) {
		if(xssfRow.getCellType()==xssfRow.CELL_TYPE_BOOLEAN) {
//			return String.valueOf(xssfRow.getBooleanCellValue());
			return xssfRow.getBooleanCellValue();
		}else if (xssfRow.getCellType()==xssfRow.CELL_TYPE_NUMERIC) {
//			return String.valueOf(xssfRow.getNumericCellValue());
			return xssfRow.getNumericCellValue();
		}else {
//			return String.valueOf(xssfRow.getStringCellValue());
			return xssfRow.getStringCellValue();
			
		}
	}

	//获取
}
