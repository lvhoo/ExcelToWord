package Bnu;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import Excel.ReadDirFiles;

public class YjbgToExcel {

	public static void main(String[] args) throws Exception {
		System.out.println("====================================研究报告Excel整理begin================================================");

		System.out.println("========================获取文件夹中的文件begin============");
		String dirPath = "E:/wuyu/研究报告";
		ReadDirFiles readDirFiles = new ReadDirFiles();
		ArrayList filePathList = readDirFiles.getFileList(dirPath);
		System.out.println("========================获取文件夹中的文件end============文件个数："+ filePathList.size());

		// 写入的Excel
		WritableWorkbook book = null;
		WritableSheet sheet = null;
		String path = "E:\\wuyu\\研究报告.xls";
		book = Workbook.createWorkbook(new File(path));
		sheet = book.createSheet("研究报告", 0);
		int rowIndex = 1;

		String filePath = "";
		String filePath2 = "";
		for (int filePathIndex = 0; filePathIndex < filePathList.size(); filePathIndex++) {
			filePath = (String) filePathList.get(filePathIndex);
			filePath2 = filePath.replaceAll("\\\\", "\\/");

			InputStream is = new FileInputStream(filePath2); // File is = new
			// File("D:/test.xls");
			Workbook rwb = Workbook.getWorkbook(is);
			Sheet rs = (Sheet) rwb.getSheet(0);
			int columnNum = rs.getColumns();
			int rowNum = rs.getRows();
			System.out.println("第" + filePathIndex + "个 Excel：" + filePath2+ ".中行列为：" + rowNum + ", " + columnNum);
			// columnNum = 30; //研究报告中的特定值

			Cell cell = null;
			String contentTemp = "";
			for (int i = 6; i <= rowNum - 1; i++) {
				for (int j = 0; j <= columnNum - 1; j++) {
					cell = rs.getCell(j, i);
					contentTemp = cell.getContents();
					if (j == 0 && contentTemp.equals("")&& contentTemp.trim().equals("")) {
						break; // 如果第一列数据为空，说明数据不存在。“论文”中名称是必须的
					}
					// System.out.print(j+","+i+":"+contentTemp+" ");

					// 写入Excel
					Label label = new Label(j, rowIndex, contentTemp);
					sheet.addCell(label);
				}
				rowIndex++;
			}
			rwb.close();
		}

		book.write();
		book.close();
		System.out.println("研究报告写入结束");
	}
}
