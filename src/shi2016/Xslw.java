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

public class Xslw {

	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> xslwMap = new HashMap<String, List<DataVo>>(); // 学术论文
		DataVo temp = null;

		InputStream xslwIs = new FileInputStream(path);
		Workbook xslwRwb = Workbook.getWorkbook(xslwIs);
		Sheet xslwRs = (Sheet) xslwRwb.getSheet(0);
		for (int xslwRowIndex = 0; xslwRowIndex < xslwRs.getRows(); xslwRowIndex++) {
			// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
			String value1 = xslwRs.getCell(0, xslwRowIndex).getContents(); // 学部作者
			String value2 = xslwRs.getCell(1, xslwRowIndex).getContents(); // 所在单位
			String value3 = xslwRs.getCell(2, xslwRowIndex).getContents(); // 通讯作者
			String value4 = xslwRs.getCell(3, xslwRowIndex).getContents();// 刊名称
			String value5 = xslwRs.getCell(4, xslwRowIndex).getContents();// 年/期
			String value6 = xslwRs.getCell(5, xslwRowIndex).getContents();// 所属学术机构
			String value7 = xslwRs.getCell(6, xslwRowIndex).getContents();// 成果资助
			String value8 = xslwRs.getCell(7, xslwRowIndex).getContents();// 备注（是否属于CSSCI）

			String value9 = xslwRs.getCell(8, xslwRowIndex).getContents();// 备注（是否属于CSSCI）
			String value10 = xslwRs.getCell(9, xslwRowIndex).getContents();// 备注（是否属于CSSCI）
			String value11 = xslwRs.getCell(10, xslwRowIndex).getContents();// 备注（是否属于CSSCI）
			String value12 = xslwRs.getCell(11, xslwRowIndex).getContents();// 备注（是否属于CSSCI）
			String value13 = xslwRs.getCell(12, xslwRowIndex).getContents();// 备注（是否属于CSSCI）

			if (value4 != null && !value4.trim().equals("")) {
				List<DataVo> tempList = null;
				if (xslwMap.containsKey(value4)) {
					tempList = xslwMap.get(value4);
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
				temp.setData9(value9);
				temp.setData10(value10);
				temp.setData11(value11);
				temp.setData12(value12);
				temp.setData13(value13);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				xslwMap.put(value4, tempList);
			}
		}
		xslwRwb.close();
		xslwIs.close();
		if (xslwMap == null || xslwMap.size() <= 0) {
		} else {
			System.out.println("============学术论文====获取成功！！！====总数为" + xslwMap.size());
		}

		return xslwMap;
	}

	public static void main(String[] args) throws Exception {
		Map<String, List<DataVo>> dataMap = Xslw.getData(true, "E:/bnu/2016/2016年科研/论文.xls");
		
	}

}
