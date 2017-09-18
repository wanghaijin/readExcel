package readExcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//工具类
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

	// 获取Excel 97-2007当前行对应列的值,并将其变为String类型
	public static String getValue(HSSFCell hssfRow) {
		if (hssfRow.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(hssfRow.getBooleanCellValue());
//			return hssfRow.getBooleanCellValue();
		}else if(hssfRow.getCellType()==hssfRow.CELL_TYPE_NUMERIC){
			//将科学计数法整数转换为源数字的字符串格式
			DecimalFormat formatNum=new DecimalFormat("#");
			return formatNum.format(hssfRow.getNumericCellValue());
//			return String.valueOf(hssfRow.getNumericCellValue()+" ");
//			return hssfRow.getNumericCellValue();
		}else {
			return String.valueOf(hssfRow.getRichStringCellValue());
//			return hssfCell.getRichStringCellValue();
		}
	}
	
	//获取Excel 2010-当前行对应列的值,并将其变为String类型
	public static String getValue(XSSFCell xssfRow) {
		if(xssfRow.getCellType()==xssfRow.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xssfRow.getBooleanCellValue());
//			return xssfRow.getBooleanCellValue();
		}else if (xssfRow.getCellType()==xssfRow.CELL_TYPE_NUMERIC) {
			DecimalFormat formatNum=new DecimalFormat("#");
			return formatNum.format(xssfRow.getNumericCellValue());
//			return String.valueOf(xssfRow.getNumericCellValue()+" ");
//			return xssfRow.getNumericCellValue();
		}else {
			return String.valueOf(xssfRow.getStringCellValue());
//			return xssfRow.getStringCellValue();
			
		}
	}


}
