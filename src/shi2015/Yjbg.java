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

public class Yjbg {

    public static Map<String, List<DataVo>> getData(boolean out) throws Exception {
	Map<String, List<DataVo>> yjbgMap = new HashMap<String, List<DataVo>>(); // 研究咨询报告
	DataVo temp = null;

	InputStream yjbgIs = new FileInputStream("E:/bnu/2015/2015年科研/2015年报告.xls");
	Workbook yjbgRwb = Workbook.getWorkbook(yjbgIs);
	Sheet yjbgRs = (Sheet) yjbgRwb.getSheet(0);
	for (int yjbgRowIndex = 1; yjbgRowIndex < yjbgRs.getRows() - 1; yjbgRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String yjbgName = yjbgRs.getCell(0, yjbgRowIndex).getContents(); // 姓名
	    String yjbgId = yjbgRs.getCell(1, yjbgRowIndex).getContents(); // ID

	    String temp1 = yjbgRs.getCell(2, yjbgRowIndex).getContents(); // 全部作者
	    String temp3 = yjbgRs.getCell(3, yjbgRowIndex).getContents(); // CGMC<成果名称>
	    String temp4 = yjbgRs.getCell(4, yjbgRowIndex).getContents(); // CNDW<采纳单位>

	    if (yjbgName != null && !yjbgName.trim().equals("")) {
		List<DataVo> tempList = null;
		if (yjbgMap.containsKey(yjbgName)) {
		    tempList = yjbgMap.get(yjbgName);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		temp.setData1(yjbgId);
		temp.setData2(yjbgName);
		temp.setData3(temp3);
		temp.setData4(temp4);
		
		if(out) {
		    System.out.println(temp.toString());
		}
		
		tempList.add(temp);
		yjbgMap.put(yjbgName, tempList);
	    }
	}
	yjbgRwb.close();
	yjbgIs.close();
	if (yjbgMap == null || yjbgMap.size() <= 0) {
	} else {
	    System.out.println("============研究咨询报告====获取成功！！！====总数为" + yjbgMap.size());
	}

	return yjbgMap;
    }

    public static void main(String[] args) throws Exception {
	Map<String, List<DataVo>> dataMap = Yjbg.getData(true);
    }

}
