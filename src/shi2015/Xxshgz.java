package shi2015;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;

public class Xxshgz {
    public static Map<String, List<DataVo>> getData(boolean out) throws Exception {
	Map<String, List<DataVo>> xxshgzMap = new HashMap<String, List<DataVo>>(); // 担任学校社会工作
	DataVo temp = null;

	InputStream xxshgzIs = new FileInputStream("E:/bnu/2015/2015社会服务/2015年度考核 担任社会工作情况（整理好）.xls");
	Workbook xxshgzRwb = Workbook.getWorkbook(xxshgzIs);
	Sheet xxshgzRs = (Sheet) xxshgzRwb.getSheet(0);
	for (int xxshgzRowIndex = 0; xxshgzRowIndex < xxshgzRs.getRows() - 1; xxshgzRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String xxshgzName1 = xxshgzRs.getCell(0, xxshgzRowIndex).getContents().trim(); // 第1完成人姓名
	    // 导出项目
	    String temp1 = xxshgzRs.getCell(1, xxshgzRowIndex).getContents();
	    if (xxshgzName1 != null && !xxshgzName1.trim().equals("")) {
		List<DataVo> tempList = null;
		if (xxshgzMap.containsKey(xxshgzName1)) {
		    tempList = xxshgzMap.get(xxshgzName1);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		temp.setData1(temp1);

		if(out) {
		    System.out.println(temp.toString());
		}
		
		tempList.add(temp);
		xxshgzMap.put(xxshgzName1, tempList);
	    }
	}
	xxshgzRwb.close();
	xxshgzIs.close();
	if (xxshgzMap == null || xxshgzMap.size() <= 0) {
	} else {
	    System.out.println("============担任学校社会工作====获取成功！！！====总数为" + xxshgzMap.size());
	}

	return xxshgzMap;
    }

    public static void main(String[] args) throws Exception {
	Map<String, List<DataVo>> dataMap = Xxshgz.getData(true);
    }

}
