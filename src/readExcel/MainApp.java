package readExcel;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class MainApp {
	public static void main(String[] args) throws IOException {
		String excel_2010="src//readExcel//≤‚ ‘Œƒµµ.xls";
//		String excel_2010="src//readExcel//ls1.xlsx";
		List<Object[][]> lists=(List<Object[][]>)new ReadExcel().readExcel(excel_2010);
		if (!lists.isEmpty()) {
//			System.out.println(lists);
			for(Object[][] strings:lists) {
				for (int i = 0; i < strings.length; i++) {
					for (int j = 0; j < strings[0].length; j++) {
						System.out.print(strings[i][j]+" ");
					}
					System.out.println();
				}
			}
		}
	}

}
