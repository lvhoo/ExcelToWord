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

public class Dsxs {

	public static Map<String, List<DataVo>> getData(boolean out, String path) throws Exception {
		Map<String, List<DataVo>> dsxsMap = new HashMap<String, List<DataVo>>(); // 导师带学生
		DataVo temp = null;
		InputStream dsxsIs = new FileInputStream(path);
		Workbook dsxsRwb = Workbook.getWorkbook(dsxsIs);
		Sheet dsxsRs = (Sheet) dsxsRwb.getSheet(0);
		for (int dsxsRowIndex = 0; dsxsRowIndex < dsxsRs.getRows(); dsxsRowIndex++) {
			// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
			String dsxsName1 = dsxsRs.getCell(0, dsxsRowIndex).getContents().trim(); // 第1完成人姓名.A
			String dsxsId1 = dsxsRs.getCell(1, dsxsRowIndex).getContents().trim(); // 第1完成人ID.B

			// 导出项目
			String temp1 = dsxsRs.getCell(2, dsxsRowIndex).getContents(); // 2011级学术型硕士C
			String temp2 = dsxsRs.getCell(3, dsxsRowIndex).getContents(); // 2011级学术型博士
			String temp3 = dsxsRs.getCell(4, dsxsRowIndex).getContents(); // 2012级学术型硕士
			String temp4 = dsxsRs.getCell(5, dsxsRowIndex).getContents(); // 2012级学术型博士
			String temp5 = dsxsRs.getCell(6, dsxsRowIndex).getContents(); // 2013级学术型硕士
			String temp6 = dsxsRs.getCell(7, dsxsRowIndex).getContents(); // 2013级学术型博士
			String temp7 = dsxsRs.getCell(8, dsxsRowIndex).getContents(); // 2011级在职硕士
			String temp8 = dsxsRs.getCell(9, dsxsRowIndex).getContents(); // 2012级在职硕士
			String temp9 = dsxsRs.getCell(10, dsxsRowIndex).getContents(); // 2013级在职硕士
			String temp10 = dsxsRs.getCell(11, dsxsRowIndex).getContents(); // 2012级免费师范生教育硕士
			String temp11 = dsxsRs.getCell(12, dsxsRowIndex).getContents(); // 2013级免费师范生教育硕士
			String temp12 = dsxsRs.getCell(13, dsxsRowIndex).getContents(); // 2011级英文国际硕士
			String temp13 = dsxsRs.getCell(14, dsxsRowIndex).getContents(); // 2012级英文国际硕士
			String temp14 = dsxsRs.getCell(15, dsxsRowIndex).getContents(); // 2013级英文国际硕士
			String temp15 = dsxsRs.getCell(16, dsxsRowIndex).getContents(); // 2013级英文国际博士

			if (dsxsName1 != null && !dsxsName1.trim().equals("")) {
				List<DataVo> tempList = null;
				if (dsxsMap.containsKey(dsxsName1)) {
					tempList = dsxsMap.get(dsxsName1);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData1(dsxsId1);
				temp.setData2(temp1);
				temp.setData3(temp2);
				temp.setData4(temp3);
				temp.setData5(temp4);
				temp.setData6(temp5);
				temp.setData7(temp6);
				temp.setData8(temp7);
				temp.setData9(temp8);
				temp.setData10(temp9);
				temp.setData11(temp10);
				temp.setData12(temp11);
				temp.setData13(temp12);
				temp.setData14(temp13);
				temp.setData15(temp14);
				temp.setData16(temp15);

				if (out) {
					System.out.println(temp.toString());
				}

				tempList.add(temp);
				dsxsMap.put(dsxsName1, tempList);
			}
		}
		dsxsRwb.close();
		dsxsIs.close();
		if (dsxsMap == null || dsxsMap.size() <= 0) {
		} else {
			System.out.println("============导师带学生====获取成功！！！====总数为" + dsxsMap.size());
		}

		return dsxsMap;
	}

	public static void main(String[] args) throws Exception {
		Map<String, List<DataVo>> dataMap = Dsxs.getData(true, "E:/bnu/2015/2015课程/学生名册.xls");
	}

}
