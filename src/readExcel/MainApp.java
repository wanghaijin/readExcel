package readExcel;

import java.io.IOException;
import java.util.List;

public class MainApp {
	public static void main(String[] args) throws IOException {
//		String excel_2003="src//main//java//excel//≤‚ ‘Œƒµµ.xls";
		String excel_2010="src//readExcel//ls1.xlsx";
		List<List<List<Object>>> listss=new ReadExcel().readExcel(excel_2010);
		if (!listss.isEmpty()) {
			for(List<List<Object>> lists:listss) 
			for(List<Object> list:lists) {
				for(Object string:list) {
					System.out.print(string+ " ");
				}
				System.out.println();
			}
		}
//		System.out.println("======================");
//		List<Model> list2=new ReadExcel().readExcel(excel_2010);
//		if (!list2.isEmpty()) {
//			for(Model model:list2) {
//				System.out.println(model);
//			}
//			
//		}
	}

}
