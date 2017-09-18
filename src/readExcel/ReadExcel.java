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
	

	// ��ȡExcel 97-2007
	private static Object readXls(String path) throws IOException {
		InputStream iStream = new FileInputStream(path);
		//Excel�ļ�
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(iStream);
		
		//�洢Excel
		List<Object[][]> list=new ArrayList<>();
		//��ȡsheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			//�洢sheet����
			List<List<Object>> lists=new ArrayList<List<Object>>();
			if (hssfSheet == null) {
				continue;
			}
			//����,�˺������ص���ȥ����һ�к������,��һ��Ϊȫ������
			int maxRowNum=hssfSheet.getLastRowNum()+1;
			//����
			int maxCellNum=hssfSheet.getRow(0).getLastCellNum();
			//�洢һ��������
			String[][] strings=new String[maxRowNum][maxCellNum];
			for (int rowNum = 0; rowNum < maxRowNum; rowNum++) {
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);

				if (hssfRow != null) {
					//��ȡÿ�е�����
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
	
	//��ȡExcel2010-����
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
			//����,�˺������ص���ȥ����һ�к������,��һ��Ϊȫ������
			int maxRowNum=xssfSheet.getLastRowNum()+1;
			//����
			int maxCellNum=xssfSheet.getRow(0).getLastCellNum();
			String [][] strings=new String[maxRowNum][maxCellNum];;
			for(int rowNum=0;rowNum<maxRowNum;rowNum++) {
				XSSFRow xssfRow=xssfSheet.getRow(rowNum);
				//��ȡ����
				maxCellNum=xssfRow.getLastCellNum();
				if (xssfRow!=null) {
					//��ȡ������,�������ݵ���list
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
