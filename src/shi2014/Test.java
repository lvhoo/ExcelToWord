package shi2014;

import java.io.FileInputStream;
import java.io.InputStream;

import com.sun.xml.internal.ws.util.StringUtils;

import jxl.Sheet;
import jxl.Workbook;

public class Test {

	public static void main(String[] args) throws Exception {

		InputStream is = new FileInputStream("D:/test/1.xls");
		Workbook rwb = Workbook.getWorkbook(is);

		Sheet[] allSheetArray = rwb.getSheets();
		//System.out.println("allSheetArray size=" + allSheetArray.length);

		Sheet sheet = null;
		String sheetName = null;
		String word = null;
		for (int sheetIndex = 0; sheetIndex < allSheetArray.length; sheetIndex++) {
			sheet = (Sheet) rwb.getSheet(sheetIndex);
			sheetName = sheet.getName();
			int rowNum = sheet.getRows();
			if (rowNum > 0) {
				for (int rowIndex = 0; rowIndex < rowNum; rowIndex++) {
					word = sheet.getCell(0,rowIndex).getContents().trim();
					if(word != null && !word.equals("")) {
						System.out.println(sheetName + "==" + word);
					}
				}
			} else {
				System.out.println(sheetName);
			}
		}
		//System.out.println("end");
	}
}
