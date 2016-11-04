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

public class Kcxx {

	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> kcxxMap = new HashMap<String, List<DataVo>>(); // 课程信息
		DataVo temp = null;
		InputStream kcxxIs = new FileInputStream(path);
		Workbook kcxxRwb = Workbook.getWorkbook(kcxxIs);
		Sheet kcxxRs = (Sheet) kcxxRwb.getSheet(0); // 区别
		for (int kcxxRowIndex = 0; kcxxRowIndex < kcxxRs.getRows(); kcxxRowIndex++) {
			// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
			String kcxxName1 = kcxxRs.getCell(0, kcxxRowIndex).getContents().trim(); // 第1完成人姓名A
			String mcTemp = kcxxRs.getCell(3, kcxxRowIndex).getContents(); // 课程名称D
			String xfTemp = kcxxRs.getCell(4, kcxxRowIndex).getContents(); // 学分E
			String bhTemp = kcxxRs.getCell(5, kcxxRowIndex).getContents(); // 上课班号F
			String xqTemp = kcxxRs.getCell(14, kcxxRowIndex).getContents(); // 学期0

			if (kcxxName1 != null && !kcxxName1.trim().equals("")) {
				List<DataVo> tempList = null;
				if (kcxxMap.containsKey(kcxxName1)) {
					tempList = kcxxMap.get(kcxxName1);
				} else {
					tempList = new ArrayList<DataVo>();
				}
				temp = new DataVo();
				temp.setData1(mcTemp);
				temp.setData2(xfTemp);
				temp.setData3(bhTemp);
				temp.setData4(xqTemp);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				kcxxMap.put(kcxxName1, tempList);
			}
		}
		kcxxRwb.close();
		kcxxIs.close();
		if (kcxxMap == null || kcxxMap.size() <= 0) {
		} else {
			System.out.println("============课程信息====获取成功！！！====总数为" + kcxxMap.size());
		}

		return kcxxMap;
	}

	public static void main(String[] args) throws Exception {
		Map<String, List<DataVo>> dataMap = Kcxx.getData(true, "E:/bnu/2015/2015课程/2015全部授课及实习.xls");
	}

}
