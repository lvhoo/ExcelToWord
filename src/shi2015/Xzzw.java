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

public class Xzzw {
    public static Map<String, List<DataVo>> getData(boolean out) throws Exception {
	Map<String, List<DataVo>> xzzwMap = new HashMap<String, List<DataVo>>(); // 行政职务
	DataVo temp = null;
	
	InputStream xzzwIs = new FileInputStream("E:/bnu/2015/2015社会服务/2015年现任行政职务.xls");
	Workbook xzzwRwb = Workbook.getWorkbook(xzzwIs);
	Sheet xzzwRs = (Sheet) xzzwRwb.getSheet(0);
	for (int xzzwRowIndex = 1; xzzwRowIndex < xzzwRs.getRows() - 1; xzzwRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String xzzwName1 = xzzwRs.getCell(0, xzzwRowIndex).getContents().trim(); // 第1完成人姓名
	    // String xzzwId1 = xzzwRs.getCell(1,
	    // xzzwRowIndex).getContents().trim(); // 第1完成人ID

	    String zwTemp = xzzwRs.getCell(1, xzzwRowIndex).getContents(); // 职务
	    String qtTemp = xzzwRs.getCell(2, xzzwRowIndex).getContents(); // 其他工作

	    if (xzzwName1 != null && !xzzwName1.trim().equals("")) {
		List<DataVo> tempList = null;
		if (xzzwMap.containsKey(xzzwName1)) {
		    tempList = xzzwMap.get(xzzwName1);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		// temp.setData1(xzzwId1);
		temp.setData2(zwTemp);
		temp.setData3(qtTemp);

		if(out) {
		    System.out.println(temp.toString());
		}
		
		tempList.add(temp);
		xzzwMap.put(xzzwName1, tempList);
	    }
	}
	xzzwRwb.close();
	xzzwIs.close();
	if (xzzwMap == null || xzzwMap.size() <= 0) {
	} else {
	    System.out.println("============行政职务====获取成功！！！====总数为" + xzzwMap.size());
	}

	return xzzwMap;
    }

    public static void main(String[] args) throws Exception {
	Map<String, List<DataVo>> dataMap = Xzzw.getData(true);
    }

}
