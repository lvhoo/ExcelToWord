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

public class Shfw {

	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> shfwMap = new HashMap<String, List<DataVo>>(); // 其他工作
		DataVo temp = null;

		InputStream shfwIs = new FileInputStream(path);
		Workbook shfwRwb = Workbook.getWorkbook(shfwIs);
		Sheet shfwRs = (Sheet) shfwRwb.getSheet(0);
		for (int shfwRowIndex = 0; shfwRowIndex < shfwRs.getRows(); shfwRowIndex++) {
			// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
			String shfwName1 = shfwRs.getCell(0, shfwRowIndex).getContents().trim(); // 第1完成人姓名
			// String shfwId1 = shfwRs.getCell(1,
			// shfwRowIndex).getContents().trim(); // 第1完成人ID

			// 导出项目
			String dwTemp = shfwRs.getCell(1, shfwRowIndex).getContents(); //
			String zwTemp = shfwRs.getCell(2, shfwRowIndex).getContents(); // 职务
			String value3 = shfwRs.getCell(3, shfwRowIndex).getContents(); // 职务

			if (shfwName1 != null && !shfwName1.trim().equals("")) {
				List<DataVo> tempList = null;
				if (shfwMap.containsKey(shfwName1)) {
					tempList = shfwMap.get(shfwName1);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData1(shfwName1);
				temp.setData2(dwTemp);
				temp.setData3(zwTemp);
				temp.setData4(value3);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				shfwMap.put(shfwName1, tempList);
			}
		}
		shfwRwb.close();
		shfwIs.close();
		if (shfwMap == null || shfwMap.size() <= 0) {
		} else {
			System.out.println("============其他工作====获取成功！！！====总数为" + shfwMap.size());
		}

		return shfwMap;
	}

	public static void main(String[] args) throws Exception {
		Map<String, List<DataVo>> dataMap = Shfw.getData(true,
				"E:/bnu/2015/2015社会服务/2015其他工作完成情况.xls");
	}

}
