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

public class Kcxx4 {

	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> kcxxMap = new HashMap<String, List<DataVo>>(); // 课程信息
		DataVo temp = null;
		InputStream kcxxIs = new FileInputStream(path);
		Workbook kcxxRwb = Workbook.getWorkbook(kcxxIs);
		Sheet kcxxRs = (Sheet) kcxxRwb.getSheet(0); // 区别
		for (int kcxxRowIndex = 0; kcxxRowIndex < kcxxRs.getRows(); kcxxRowIndex++) {
			String daoshiAll = kcxxRs.getCell(0, kcxxRowIndex).getContents().trim();
			String kecheng = kcxxRs.getCell(3, kcxxRowIndex).getContents();
			String banhao = kcxxRs.getCell(5, kcxxRowIndex).getContents();
			String shijian = kcxxRs.getCell(14, kcxxRowIndex).getContents();

			if (daoshiAll != null && !daoshiAll.trim().equals("")) {
				String[] daoshiArray = daoshiAll.split(" ");
				if(daoshiArray != null && daoshiArray.length > 0) {
					for(int i=0; i<daoshiArray.length; i++) {
						String daoshi = daoshiArray[i].trim();
						// System.out.println(daoshi);
						List<DataVo> tempList = null;
						if (kcxxMap.containsKey(daoshi)) {
							tempList = kcxxMap.get(daoshi);
						} else {
							tempList = new ArrayList<DataVo>();
						}
						temp = new DataVo();
						temp.setData1(kecheng);
						temp.setData2(banhao);
						temp.setData3(shijian);

						if (out) {
							System.out.println(temp.toString());
						}

						tempList.add(temp);
						kcxxMap.put(daoshi, tempList);
					}
				}
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
		Map<String, List<DataVo>> dataMap = Kcxx4.getData(true, "E:/bnu/2016/实习合并（给田桦）.xls");
	}

}
