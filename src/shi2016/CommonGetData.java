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

public class CommonGetData {

	public static Map<String, List<DataVo>> getData(String filePath, int index, boolean out)
			throws Exception {
		Map<String, List<DataVo>> dataMap = new HashMap<String, List<DataVo>>();
		DataVo temp = null;

		System.out.println(filePath + "============begin===================");
		InputStream xslwIs = new FileInputStream(filePath);
		Workbook xslwRwb = Workbook.getWorkbook(xslwIs);
		Sheet xslwRs = (Sheet) xslwRwb.getSheet(0);
		for (int xslwRowIndex = 0; xslwRowIndex < xslwRs.getRows() - 1; xslwRowIndex++) {
			String value0 = xslwRs.getCell(0, xslwRowIndex).getContents();
			String value1 = xslwRs.getCell(1, xslwRowIndex).getContents();
			String value2 = xslwRs.getCell(2, xslwRowIndex).getContents();
			String value3 = xslwRs.getCell(3, xslwRowIndex).getContents();
			String value4 = xslwRs.getCell(4, xslwRowIndex).getContents();
			String value5 = xslwRs.getCell(5, xslwRowIndex).getContents();
			String value6 = xslwRs.getCell(6, xslwRowIndex).getContents();
			String value7 = xslwRs.getCell(7, xslwRowIndex).getContents();
			String value8 = xslwRs.getCell(8, xslwRowIndex).getContents();
			String value9 = xslwRs.getCell(9, xslwRowIndex).getContents();
			String value10 = xslwRs.getCell(10, xslwRowIndex).getContents();
			String value11 = xslwRs.getCell(11, xslwRowIndex).getContents();
			String value12 = xslwRs.getCell(12, xslwRowIndex).getContents();
			String value13 = xslwRs.getCell(13, xslwRowIndex).getContents();
			String value14 = xslwRs.getCell(14, xslwRowIndex).getContents();
			String value15 = xslwRs.getCell(15, xslwRowIndex).getContents();
			String value16 = xslwRs.getCell(16, xslwRowIndex).getContents();
			String value17 = xslwRs.getCell(17, xslwRowIndex).getContents();
			String value18 = xslwRs.getCell(18, xslwRowIndex).getContents();
			String value19 = xslwRs.getCell(19, xslwRowIndex).getContents();
			String value20 = xslwRs.getCell(20, xslwRowIndex).getContents();

			String key = value0;
			switch (index) {
			case 0:
				key = value0;
				break;
			case 1:
				key = value1;
				break;
			case 2:
				key = value2;
				break;
			case 3:
				key = value3;
				break;
			case 4:
				key = value4;
				break;
			case 5:
				key = value5;
				break;
			case 6:
				key = value6;
				break;
			case 7:
				key = value7;
				break;
			case 8:
				key = value8;
				break;
			case 9:
				key = value9;
				break;
			case 10:
				key = value10;
				break;
			case 11:
				key = value11;
				break;
			case 12:
				key = value12;
				break;
			case 13:
				key = value13;
				break;
			case 14:
				key = value14;
				break;
			case 15:
				key = value15;
				break;
			case 16:
				key = value16;
				break;
			case 17:
				key = value17;
				break;
			case 18:
				key = value18;
				break;
			case 19:
				key = value19;
				break;
			case 20:
				key = value20;
				break;
			}

			if (key != null && !key.trim().equals("")) {
				List<DataVo> tempList = null;
				if (dataMap.containsKey(key)) {
					tempList = dataMap.get(key);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData0(value0);
				temp.setData1(value1);
				temp.setData2(value2);
				temp.setData3(value3);
				temp.setData4(value4);
				temp.setData5(value5);
				temp.setData6(value6);
				temp.setData7(value7);
				temp.setData8(value8);
				temp.setData9(value9);
				temp.setData10(value10);
				temp.setData11(value11);
				temp.setData12(value12);
				temp.setData13(value13);
				temp.setData14(value14);
				temp.setData15(value15);
				temp.setData16(value16);
				temp.setData17(value17);
				temp.setData18(value18);
				temp.setData19(value19);
				temp.setData20(value20);

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

		return dataMap;
	}

	public static void main(String[] args) throws Exception {
	}

}
