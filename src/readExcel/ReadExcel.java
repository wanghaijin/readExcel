package readExcel;

import java.io.IOException;
import java.util.List;

public class ReadExcel {
	
	public List<List<List<Object>>> readExcel(String path) throws IOException {
		if(path == null || Common.EMPTY.equals(path)) {
			return null;
		}else {
			String postfix=Util.getPostfix(path);
			if(!Common.EMPTY.equals(postfix)) {
				if(Common.OFFICE_2003.equals(postfix)) {
					return Util.readXls(path);
				}else if (Common.OFFICE_2010.equals(postfix)) {
					return Util.readXlsx(path);
				}else {
					System.out.println(path+" "+ Common.NOT_EXCEL_FILE);
				}
			}
			return null;
		}
		
	}

}
