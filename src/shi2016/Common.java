package shi2016;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;

public class Common {

	public static Map<String, List<DataVo>> getData(String filePath, int index, boolean out)
			throws Exception {
		Map<String, List<DataVo>> dataMap = new HashMap<String, List<DataVo>>();
		DataVo temp = null;

		System.out.println(filePath + "============begin===================");
		InputStream xslwIs = new FileInputStream(filePath);
		Workbook xslwRwb = Workbook.getWorkbook(xslwIs);
		Sheet xslwRs = (Sheet) xslwRwb.getSheet(0);
		for (int xslwRowIndex = 0; xslwRowIndex < xslwRs.getRows() - 1; xslwRowIndex++) {
			String value1 = xslwRs.getCell(0, xslwRowIndex).getContents();
			String value2 = xslwRs.getCell(1, xslwRowIndex).getContents();
			String value3 = xslwRs.getCell(2, xslwRowIndex).getContents();
			String value4 = xslwRs.getCell(3, xslwRowIndex).getContents();
			String value5 = xslwRs.getCell(4, xslwRowIndex).getContents();
			String value6 = xslwRs.getCell(5, xslwRowIndex).getContents();
			String value7 = xslwRs.getCell(6, xslwRowIndex).getContents();
			String value8 = xslwRs.getCell(7, xslwRowIndex).getContents();
			String value9 = xslwRs.getCell(8, xslwRowIndex).getContents();
			String value10 = xslwRs.getCell(9, xslwRowIndex).getContents();
			String value11 = xslwRs.getCell(10, xslwRowIndex).getContents();
			String value12 = xslwRs.getCell(11, xslwRowIndex).getContents();
			String value13 = xslwRs.getCell(12, xslwRowIndex).getContents();
			String value14 = xslwRs.getCell(13, xslwRowIndex).getContents();
			String value15 = xslwRs.getCell(14, xslwRowIndex).getContents();

			String key = value1;
			if (index == 1) {
				key = value1;
			} else if (index == 2) {
				key = value2;
			} else if (index == 3) {
				key = value3;
			} else if (index == 4) {
				key = value4;
			} else if (index == 5) {
				key = value5;
			} else if (index == 6) {
				key = value6;
			} else if (index == 7) {
				key = value7;
			} else if (index == 8) {
				key = value8;
			} else if (index == 9) {
				key = value9;
			} else if (index == 10) {
				key = value10;
			} else if (index == 11) {
				key = value11;
			} else if (index == 12) {
				key = value12;
			} else if (index == 13) {
				key = value13;
			} else if (index == 14) {
				key = value14;
			} else if (index == 15) {
				key = value15;
			}

			if (key != null && !key.trim().equals("")) {
				List<DataVo> tempList = null;
				if (dataMap.containsKey(key)) {
					tempList = dataMap.get(key);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData1(value1);
				temp.setData2(value2);
				temp.setData3(value3);
				temp.setData4(value4);
				temp.setData5(value5);
				temp.setData6(value6);
				temp.setData7(value7);
				temp.setData8(value8);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				dataMap.put(key, tempList);
			}
		}
		xslwRwb.close();
		xslwIs.close();
		if (dataMap == null || dataMap.size() <= 0) {
			System.out.println("为空！！！！");
		} else {
			System.out.println("获取成功！！！====总数为" + dataMap.size());
		}
		System.out.println("end==");

		return dataMap;
	}

	public static void main(String[] args) throws Exception {
	}

}
