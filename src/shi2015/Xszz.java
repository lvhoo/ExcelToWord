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

public class Xszz {

    public static Map<String, List<DataVo>> getData(boolean out) throws Exception {
	Map<String, List<DataVo>> xszzMap = new HashMap<String, List<DataVo>>(); // 学术著作
	DataVo temp = null;

	InputStream zzIs = new FileInputStream("E:/bnu/2015/2015年科研/2015年著作.xls");
	Workbook zzRwb = Workbook.getWorkbook(zzIs);
	Sheet zzRs = (Sheet) zzRwb.getSheet(0);
	for (int zzRowIndex = 1; zzRowIndex < zzRs.getRows() - 1; zzRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String temp1 = zzRs.getCell(0, zzRowIndex).getContents();
	    String temp2 = zzRs.getCell(1, zzRowIndex).getContents();
	    String temp3 = zzRs.getCell(2, zzRowIndex).getContents();
	    String temp4 = zzRs.getCell(3, zzRowIndex).getContents();
	    String temp5 = zzRs.getCell(4, zzRowIndex).getContents();
	    String temp6 = zzRs.getCell(5, zzRowIndex).getContents();
	    String temp7 = zzRs.getCell(6, zzRowIndex).getContents();

	    if (temp2 != null && !temp2.trim().equals("")) {
		List<DataVo> tempList = null;
		if (xszzMap.containsKey(temp2)) {
		    tempList = xszzMap.get(temp2);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		temp.setData1(temp1);
		temp.setData2(temp2);
		temp.setData3(temp3);
		temp.setData4(temp4);
		temp.setData5(temp5);
		temp.setData6(temp6);
		temp.setData7(temp7);

		if(out) {
		    System.out.println(temp.toString());
		}
		
		tempList.add(temp);
		xszzMap.put(temp2, tempList);
	    }
	}
	zzRwb.close();
	zzIs.close();
	if (xszzMap == null || xszzMap.size() <= 0) {
	} else {
	    System.out.println("============学术著作====获取成功！！！====总数为" + xszzMap.size());
	}

	return xszzMap;
    }

    public static void main(String[] args) throws Exception {
	Map<String, List<DataVo>> dataMap = Xszz.getData(true);
    }

}
