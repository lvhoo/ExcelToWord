package shi2015;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;

public class Kyxm {
    public static Map<String, List<DataVo>> getData(boolean out) throws Exception {
	Map<String, List<DataVo>> kyxmMap = new HashMap<String, List<DataVo>>(); // 科研项目
	DataVo temp = null;

	InputStream kyxmIs = new FileInputStream("E:/bnu/2015/2015年科研/2015年省部级项目.xls");
	Workbook kyxmRwb = Workbook.getWorkbook(kyxmIs);
	Sheet kyxmRs = (Sheet) kyxmRwb.getSheet(0);
	for (int kyxmRowIndex = 1; kyxmRowIndex < kyxmRs.getRows() - 1; kyxmRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String kyxmName1 = kyxmRs.getCell(0, kyxmRowIndex).getContents().trim(); // 作者
	    String kyxmId = kyxmRs.getCell(1, kyxmRowIndex).getContents(); // ID
	    String temp1 = kyxmRs.getCell(2, kyxmRowIndex).getContents(); // 项目来源
	    String temp3 = kyxmRs.getCell(4, kyxmRowIndex).getContents(); // 项目类型
	    String temp4 = kyxmRs.getCell(6, kyxmRowIndex).getContents(); // 项目名称
	    String temp5 = kyxmRs.getCell(5, kyxmRowIndex).getContents(); // 项目批准号

	    if (kyxmName1 != null && !kyxmName1.trim().equals("")) {
		List<DataVo> tempList = null;
		if (kyxmMap.containsKey(kyxmName1)) {
		    tempList = kyxmMap.get(kyxmName1);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		temp.setData1(kyxmId);
		temp.setData2(temp1);
		temp.setData3(temp3);
		temp.setData4(temp4);
		temp.setData5(temp5);

		if(out) {
		    System.out.println(temp.toString());
		}
		
		tempList.add(temp);
		kyxmMap.put(kyxmName1, tempList);
	    }
	}
	kyxmRwb.close();
	kyxmIs.close();
	if (kyxmMap == null || kyxmMap.size() <= 0) {
	} else {
	    System.out.println("============科研项目====获取成功！！！====总数为" + kyxmMap.size());
	}

	return kyxmMap;
    }

    public static void main(String[] args) throws Exception {
	Map<String, List<DataVo>> dataMap = Kyxm.getData(true);
    }

}
