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

public class Kyhj {
	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> kyhjMap = new HashMap<String, List<DataVo>>(); // 科研获奖
		DataVo temp = null;

		InputStream kyhjIs = new FileInputStream(path);
		Workbook kyhjRwb = Workbook.getWorkbook(kyhjIs);
		Sheet kyhjRs = (Sheet) kyhjRwb.getSheet(0);
		for (int kyhjRowIndex = 0; kyhjRowIndex < kyhjRs.getRows(); kyhjRowIndex++) {
			// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
			String kyhjName = kyhjRs.getCell(1, kyhjRowIndex).getContents(); // 姓名
			// String kyhjId = kyhjRs.getCell(1, kyhjRowIndex).getContents(); //
			// ID

			String temp1 = kyhjRs.getCell(2, kyhjRowIndex).getContents(); // 获奖者
			String temp2 = kyhjRs.getCell(3, kyhjRowIndex).getContents(); // 成果名称
			String temp3 = kyhjRs.getCell(4, kyhjRowIndex).getContents(); // 成果类型
			String temp4 = kyhjRs.getCell(5, kyhjRowIndex).getContents(); // 获奖等级
			String temp5 = kyhjRs.getCell(6, kyhjRowIndex).getContents(); // 获奖名称
			String temp6 = kyhjRs.getCell(7, kyhjRowIndex).getContents(); // 获奖名称

			if (kyhjName != null && !kyhjName.trim().equals("")) {
				List<DataVo> tempList = null;
				if (kyhjMap.containsKey(kyhjName)) {
					tempList = kyhjMap.get(kyhjName);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData1(kyhjName);
				temp.setData2(temp1);
				temp.setData3(temp2);
				temp.setData4(temp3);
				temp.setData5(temp4);
				temp.setData6(temp5);
				temp.setData7(temp6);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				kyhjMap.put(kyhjName, tempList);
			}
		}
		kyhjRwb.close();
		kyhjIs.close();
		if (kyhjMap == null || kyhjMap.size() <= 0) {
		} else {
			System.out.println("============科研获奖====获取成功！！！====总数为" + kyhjMap.size());
		}

		return kyhjMap;
	}

	public static void main(String[] args) throws Exception {
		Map<String, List<DataVo>> dataMap = Kyhj.getData(true, "E:/bnu/2016/2016年科研/获奖.xls");
	}

}
