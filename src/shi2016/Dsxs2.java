package shi2016;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;

public class Dsxs2 {

	public static Map<String, Map<String, List<DataVo>>> getData(boolean out, String path)
			throws Exception {
		Map<String, Map<String, List<DataVo>>> dsxsMap = new HashMap<String, Map<String, List<DataVo>>>(); // 导师带学生
		DataVo temp = null;
		InputStream dsxsIs = new FileInputStream(path);
		Workbook dsxsRwb = Workbook.getWorkbook(dsxsIs);
		Sheet dsxsRs = (Sheet) dsxsRwb.getSheet(1);
		for (int dsxsRowIndex = 0; dsxsRowIndex < dsxsRs.getRows(); dsxsRowIndex++) {
			// String temp1 = dsxsRs.getCell(0,
			// dsxsRowIndex).getContents().trim(); // 学号
			String xingming = dsxsRs.getCell(1, dsxsRowIndex).getContents().trim(); // 姓名
			String daoshiName = dsxsRs.getCell(2, dsxsRowIndex).getContents(); // 导师姓名
			String daoshiNo = dsxsRs.getCell(3, dsxsRowIndex).getContents(); // 工作证号
			String leixing = dsxsRs.getCell(4, dsxsRowIndex).getContents(); // 类型

			// System.out.println(temp1+";"+temp2+";"+dsxsName1+";"+temp4+";"+temp5);

			if (daoshiName != null && !daoshiName.trim().equals("")) {
				Map<String, List<DataVo>> dataMap = null;
				if (dsxsMap.containsKey(daoshiName)) {
					dataMap = dsxsMap.get(daoshiName);
				} else {
					dataMap = new HashMap<String, List<DataVo>>();
				}

				List<DataVo> tempList = null;
				if (dataMap.containsKey(leixing)) {
					tempList = dataMap.get(leixing);
				} else {
					tempList = new ArrayList<DataVo>();
				}

				temp = new DataVo();
				temp.setData1(xingming);
				temp.setData2(leixing);
				tempList.add(temp);

				dataMap.put(leixing, tempList);

				dsxsMap.put(daoshiName, dataMap);
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
		Map<String, Map<String, List<DataVo>>> dsxsMap = Dsxs2.getData(false, "E:/bnu/2016/学生名册汇总（给田桦）.xls");

		/*
		for (Map.Entry entry : dsxsMap.entrySet()) {
			String daoshi = entry.getKey().toString();
			Map<String, List<DataVo>> dataMap = (Map) entry.getValue();
			System.out.println(daoshi + ":");
			for (Map.Entry entry2 : dataMap.entrySet()) {
				String leixing = entry2.getKey().toString();
				List<DataVo> dataList = (List) entry2.getValue();
				System.out.print("  "+leixing + ":");
				for (DataVo temp : dataList) {
					System.out.print(temp.getData1()+",");
				}
				System.out.println();
			}
		}
		 * */
		
		String userName= "洪秀敏";
		
		List<String> baseList = new LinkedList<String>();
		baseList.add("2012级本科生");
		baseList.add("2013级本科生");
		baseList.add("2014级本科生");
		baseList.add("2015级本科生");
		baseList.add("2013级学术型硕士");
		baseList.add("2014级学术型硕士");
		baseList.add("2015级学术型硕士");
		baseList.add("2014级全日制教育硕士");
		baseList.add("2015级全日制教育硕士");
		baseList.add("2013级学术性博士");
		baseList.add("2014级学术性博士");
		baseList.add("2015级学术性博士");
		baseList.add("2014级暑期教育硕士");
		baseList.add("2015级暑期教育硕士");
		baseList.add("2014级英文国际硕士");
		baseList.add("2015级英文国际硕士");
		baseList.add("2013级英文国际博士");
		baseList.add("2014级英文国际博士");
		baseList.add("2015级英文国际博士");
		
		StringBuffer dsxsContent = new StringBuffer();
		if (dsxsMap != null && dsxsMap.containsKey(userName)) {
			Map<String, List<DataVo>> dataMap = dsxsMap.get(userName);
			if(dataMap != null) {
				int index = 1;
				for(String base : baseList) {
					if(dataMap.containsKey(base)) {
						List<DataVo> dataList = dataMap.get(base);
						if(dataList != null && dataList.size() > 0) {
							dsxsContent.append(index + ". " + base + ": ");
							for(int dsxsLoop=0; dsxsLoop<dataList.size(); dsxsLoop++) {
								DataVo temp = dataList.get(dsxsLoop);
								if(dsxsLoop == 0) {
									dsxsContent.append(temp.getData1());
								} else {
									dsxsContent.append(", "+ temp.getData1());
								}
							}
							dsxsContent.append("\n");
							index++;
						}
					}
				}
			}
		}
		
		System.out.println(dsxsContent);
		
	}

}
