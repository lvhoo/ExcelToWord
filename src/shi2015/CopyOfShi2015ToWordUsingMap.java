package shi2015;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;
import word.ModifyWordDocument;
import word.WordTable;

public class CopyOfShi2015ToWordUsingMap {

    // 教学工作量/课程信息汇总(2012年度).xls
    // 教学工作量/导师带学生汇总表.xls
    // 社会服务工作量/其他工作汇总.xls
    // 社会服务工作量/行政职务.xls
    // 社会服务工作量/担任学校社会工作.xls
    
    // 科研工作量/学术论文（终稿）.xls
    // 科研工作量/著作（终稿）.xls
    // 科研工作量/科研获奖（终稿）.xls
    // 科研工作量/研究咨询报告（终稿）.xls
    // 科研工作量/2011新立项国家、省部、纵向、境外合作科研项目（终稿）.xls

    public static void main(String[] args) throws Exception {
	List<DataVo> peopeInfoList = new ArrayList<DataVo>();
	Map<String, List<DataVo>> kcxxMap = new HashMap<String, List<DataVo>>(); // 课程信息
	Map<String, List<DataVo>> dsxsMap = new HashMap<String, List<DataVo>>(); // 导师带学生
	Map<String, List<DataVo>> shfwMap = new HashMap<String, List<DataVo>>(); // 其他工作
	Map<String, List<DataVo>> xzzwMap = new HashMap<String, List<DataVo>>(); // 行政职务
	Map<String, List<DataVo>> xxshgzMap = new HashMap<String, List<DataVo>>(); // 担任学校社会工作
	Map<String, List<DataVo>> xslwMap = new HashMap<String, List<DataVo>>(); // 学术论文
	Map<String, List<DataVo>> xszzMap = new HashMap<String, List<DataVo>>(); // 学术著作
	Map<String, List<DataVo>> kyhjMap = new HashMap<String, List<DataVo>>(); // 科研获奖
	Map<String, List<DataVo>> yjbgMap = new HashMap<String, List<DataVo>>(); // 研究咨询报告
	Map<String, List<DataVo>> kyxmMap = new HashMap<String, List<DataVo>>(); // 科研项目
	DataVo temp = null;

	System.out.println("============人员基本信息收集begin===================");
	InputStream is = new FileInputStream("D:/bnu/2015/people2015.xls");
	Workbook rwb = Workbook.getWorkbook(is);
	Sheet rs = (Sheet) rwb.getSheet(0);
	for (int i = 1; i <= rs.getRows() - 1; i++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String dept = "教育学部" + rs.getCell(0, i).getContents(); // 部门
	    String userId = rs.getCell(1, i).getContents(); // ID
	    String userName = rs.getCell(2, i).getContents(); // name
	    String title = rs.getCell(3, i).getContents(); // 职称
	    String grade = rs.getCell(4, i).getContents(); // 岗位级别
	    if (!userId.trim().equals("") && !userName.trim().equals("")) {
		temp = new DataVo();
		temp.setData1(dept); // 部门
		temp.setData2(userId); // ID
		temp.setData3(userName); // name
		temp.setData4(title); // 职称
		temp.setData5(grade); // 岗位级别

		peopeInfoList.add(temp);
	    }
	}
	rwb.close();
	is.close();
	if (peopeInfoList == null || peopeInfoList.size() <= 0) {
	    System.out.println("============人员信息====为空！！！！");
	} else {
	    System.out.println("============人员信息====获取成功！！！====总人数为=" + peopeInfoList.size());
	}
	System.out.println("============人员基本信息收集end===================");

	System.out.println("============课程信息汇总begin===================");
	InputStream kcxxIs = new FileInputStream("D:/bnu/2015/2014年教学工作量/2014全部授课及实习.xls");
	Workbook kcxxRwb = Workbook.getWorkbook(kcxxIs);
	Sheet kcxxRs = (Sheet) kcxxRwb.getSheet(0); // 区别
	for (int kcxxRowIndex = 1; kcxxRowIndex < kcxxRs.getRows() - 1; kcxxRowIndex++) {
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

		tempList.add(temp);
		kcxxMap.put(kcxxName1, tempList);
	    }
	}
	kcxxRwb.close();
	kcxxIs.close();
	if (kcxxMap == null || kcxxMap.size() <= 0) {
	    System.out.println("============课程信息====为空！！！！");
	} else {
	    System.out.println("============课程信息====获取成功！！！====总数为" + kcxxMap.size());
	}
	System.out.println("============课程信息汇总end===================");

	System.out.println("============导师带学生汇总begin===================");
	InputStream dsxsIs = new FileInputStream("D:/bnu/2015/2014年教学工作量/2011-2013年导师带学生数20131106（硕博士）全.xls");
	Workbook dsxsRwb = Workbook.getWorkbook(dsxsIs);
	Sheet dsxsRs = (Sheet) dsxsRwb.getSheet(0);
	for (int dsxsRowIndex = 1; dsxsRowIndex < dsxsRs.getRows() - 1; dsxsRowIndex++) {
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
	    String temp16 = dsxsRs.getCell(17, dsxsRowIndex).getContents(); // 2011级全日制教育硕士
	    String temp17 = dsxsRs.getCell(18, dsxsRowIndex).getContents(); // 2012级全日制教育硕士
	    String temp18 = dsxsRs.getCell(19, dsxsRowIndex).getContents(); // 2013级全日制教育硕士T

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
		temp.setData17(temp16);
		temp.setData18(temp17);
		temp.setData19(temp18);

		tempList.add(temp);
		dsxsMap.put(dsxsName1, tempList);
	    }
	}
	dsxsRwb.close();
	dsxsIs.close();
	if (dsxsMap == null || dsxsMap.size() <= 0) {
	    System.out.println("============导师带学生====为空！！！！");
	} else {
	    System.out.println("============导师带学生====获取成功！！！====总数为" + dsxsMap.size());
	}
	System.out.println("============导师带学生汇总end===================");

	System.out.println("============其他工作汇总begin===================");
	InputStream shfwIs = new FileInputStream("D:/bnu/2015/2014社会服务/2014其他工作汇总（整理好）.xls");
	Workbook shfwRwb = Workbook.getWorkbook(shfwIs);
	Sheet shfwRs = (Sheet) shfwRwb.getSheet(0);
	for (int shfwRowIndex = 1; shfwRowIndex < shfwRs.getRows() - 1; shfwRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String shfwName1 = shfwRs.getCell(0, shfwRowIndex).getContents().trim(); // 第1完成人姓名
	    // String shfwId1 = shfwRs.getCell(1,
	    // shfwRowIndex).getContents().trim(); // 第1完成人ID

	    // 导出项目
	    String dwTemp = shfwRs.getCell(1, shfwRowIndex).getContents(); //
	    String zwTemp = shfwRs.getCell(2, shfwRowIndex).getContents(); // 职务

	    if (shfwName1 != null && !shfwName1.trim().equals("")) {
		List<DataVo> tempList = null;
		if (shfwMap.containsKey(shfwName1)) {
		    tempList = shfwMap.get(shfwName1);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		// temp.setData1(shfwId1);
		temp.setData2(dwTemp);
		temp.setData3(zwTemp);

		tempList.add(temp);
		shfwMap.put(shfwName1, tempList);
	    }
	}
	shfwRwb.close();
	shfwIs.close();
	if (shfwMap == null || shfwMap.size() <= 0) {
	    System.out.println("============其他工作====为空！！！！");
	} else {
	    System.out.println("============其他工作====获取成功！！！====总数为" + shfwMap.size());
	}
	System.out.println("============其他工作汇总end===================");

	System.out.println("============行政职务begin===================");
	InputStream xzzwIs = new FileInputStream("D:/bnu/2015/2014社会服务/2014行政职务（整理好）.xls");
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

		tempList.add(temp);
		xzzwMap.put(xzzwName1, tempList);
	    }
	}
	xzzwRwb.close();
	xzzwIs.close();
	if (xzzwMap == null || xzzwMap.size() <= 0) {
	    System.out.println("============行政职务====为空！！！！");
	} else {
	    System.out.println("============行政职务====获取成功！！！====总数为" + xzzwMap.size());
	}
	System.out.println("============行政职务end===================");

	System.out.println("============担任学校社会工作begin===================");
	InputStream xxshgzIs = new FileInputStream("D:/bnu/2015/2014社会服务/2014年度考核 担任社会工作情况（整理好）.xls");
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

		tempList.add(temp);
		xxshgzMap.put(xxshgzName1, tempList);
	    }
	}
	xxshgzRwb.close();
	xxshgzIs.close();
	if (xxshgzMap == null || xxshgzMap.size() <= 0) {
	    System.out.println("============担任学校社会工作====为空！！！！");
	} else {
	    System.out.println("============担任学校社会工作====获取成功！！！====总数为" + xxshgzMap.size());
	}
	System.out.println("============担任学校社会工作end===================");

	System.out.println("============学术论文begin===================");
	InputStream xslwIs = new FileInputStream("D:/bnu/2015/2015年科研/学术论文.xls");
	Workbook xslwRwb = Workbook.getWorkbook(xslwIs);
	Sheet xslwRs = (Sheet) xslwRwb.getSheet(0);
	for (int xslwRowIndex = 0; xslwRowIndex < xslwRs.getRows() - 1; xslwRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String xslwName = xslwRs.getCell(0, xslwRowIndex).getContents(); // 姓名
	    String xslwId = xslwRs.getCell(1, xslwRowIndex).getContents(); // ID
	    String xslwPaper = xslwRs.getCell(2, xslwRowIndex).getContents(); // 文章

	    if (xslwName != null && !xslwName.trim().equals("")) {
		List<DataVo> tempList = null;
		if (xslwMap.containsKey(xslwName)) {
		    tempList = xslwMap.get(xslwName);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		temp.setData1(xslwId);
		temp.setData2(xslwPaper);

		tempList.add(temp);
		xslwMap.put(xslwName, tempList);
	    }
	}
	xslwRwb.close();
	xslwIs.close();
	if (xslwMap == null || xslwMap.size() <= 0) {
	    System.out.println("============学术论文====为空！！！！");
	} else {
	    System.out.println("============学术论文====获取成功！！！====总数为" + xslwMap.size());
	}
	System.out.println("============学术论文end===================");

	System.out.println("============学术著作begin===================");
	InputStream zzIs = new FileInputStream("D:/bnu/2015/2015年科研/专著.xls");
	Workbook zzRwb = Workbook.getWorkbook(zzIs);
	Sheet zzRs = (Sheet) zzRwb.getSheet(0);
	for (int zzRowIndex = 1; zzRowIndex < zzRs.getRows() - 1; zzRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String zzName = zzRs.getCell(0, zzRowIndex).getContents(); // 姓名
	    // String zzId = zzRs.getCell(1, zzRowIndex).getContents(); // ID

	    String temp1 = zzRs.getCell(2, zzRowIndex).getContents(); // 全部作者
	    String temp2 = zzRs.getCell(3, zzRowIndex).getContents(); // 著作名称
	    String temp3 = zzRs.getCell(4, zzRowIndex).getContents(); // 出版日期
	    String temp4 = zzRs.getCell(5, zzRowIndex).getContents(); // 出版社

	    if (zzName != null && !zzName.trim().equals("")) {
		List<DataVo> tempList = null;
		if (xszzMap.containsKey(zzName)) {
		    tempList = xszzMap.get(zzName);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		// temp.setData1(zzId);
		temp.setData2(temp1);
		temp.setData3(temp2);
		temp.setData4(temp3);
		temp.setData5(temp4);

		tempList.add(temp);
		xszzMap.put(zzName, tempList);
	    }
	}
	zzRwb.close();
	zzIs.close();
	if (xszzMap == null || xszzMap.size() <= 0) {
	    System.out.println("============学术著作====为空！！！！");
	} else {
	    System.out.println("============学术著作====获取成功！！！====总数为" + xszzMap.size());
	}
	System.out.println("============学术著作end===================");

	System.out.println("============科研获奖（终稿）begin===================");
	InputStream kyhjIs = new FileInputStream("D:/bnu/2015/2015年科研/奖励.xls");
	Workbook kyhjRwb = Workbook.getWorkbook(kyhjIs);
	Sheet kyhjRs = (Sheet) kyhjRwb.getSheet(0);
	for (int kyhjRowIndex = 1; kyhjRowIndex < kyhjRs.getRows() - 1; kyhjRowIndex++) {
	    // 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.
	    String kyhjName = kyhjRs.getCell(0, kyhjRowIndex).getContents(); // 姓名
	    // String kyhjId = kyhjRs.getCell(1, kyhjRowIndex).getContents(); //
	    // ID

	    String temp1 = kyhjRs.getCell(1, kyhjRowIndex).getContents(); // 获奖者
	    String temp2 = kyhjRs.getCell(2, kyhjRowIndex).getContents(); // 成果名称
	    String temp3 = kyhjRs.getCell(3, kyhjRowIndex).getContents(); // 成果类型
	    String temp4 = kyhjRs.getCell(4, kyhjRowIndex).getContents(); // 获奖等级
	    String temp5 = kyhjRs.getCell(5, kyhjRowIndex).getContents(); // 获奖名称

	    if (kyhjName != null && !kyhjName.trim().equals("")) {
		List<DataVo> tempList = null;
		if (kyhjMap.containsKey(kyhjName)) {
		    tempList = kyhjMap.get(kyhjName);
		} else {
		    tempList = new ArrayList<DataVo>();
		}

		temp = new DataVo();
		// temp.setData1(kyhjId);
		temp.setData2(temp1);
		temp.setData3(temp2);
		temp.setData4(temp3);
		temp.setData5(temp4);
		temp.setData6(temp5);

		tempList.add(temp);
		kyhjMap.put(kyhjName, tempList);
	    }
	}
	kyhjRwb.close();
	kyhjIs.close();
	if (kyhjMap == null || kyhjMap.size() <= 0) {
	    System.out.println("============科研获奖====为空！！！！");
	} else {
	    System.out.println("============科研获奖====获取成功！！！====总数为" + kyhjMap.size());
	}
	System.out.println("============科研获奖（终稿）end===================");

	System.out.println("============研究咨询报告（终稿）.xlsbegin===================");
	InputStream yjbgIs = new FileInputStream("D:/bnu/2015/2015年科研/咨询报告.xls");
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
		temp.setData2(temp1);
		temp.setData3(temp3);
		temp.setData4(temp4);

		tempList.add(temp);
		yjbgMap.put(yjbgName, tempList);
	    }
	}
	yjbgRwb.close();
	yjbgIs.close();
	if (yjbgMap == null || yjbgMap.size() <= 0) {
	    System.out.println("============研究咨询报告====为空！！！！");
	} else {
	    System.out.println("============研究咨询报告====获取成功！！！====总数为" + yjbgMap.size());
	}
	System.out.println("============研究咨询报告（终稿）.xlsend===================");

	System.out.println("============科研项目begin===================");
	InputStream kyxmIs = new FileInputStream("D:/bnu/2015/2015年科研/2013新立项项目（终）.xls");
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

		tempList.add(temp);
		kyxmMap.put(kyxmName1, tempList);
	    }
	}
	kyxmRwb.close();
	kyxmIs.close();
	if (kyxmMap == null || kyxmMap.size() <= 0) {
	    System.out.println("============科研项目====为空！！！！");
	} else {
	    System.out.println("============科研项目====获取成功！！！====总数为" + kyxmMap.size());
	}
	System.out.println("============科研项目end===================");

	
	
	
	if (peopeInfoList != null && peopeInfoList.size() > 0) {
	    File wordFile = null;
	    for (int i = 0; i < peopeInfoList.size(); i++) {
		DataVo userInfo = (DataVo) peopeInfoList.get(i);
		String dept = userInfo.getData1(); // 部门
		String userId = userInfo.getData2(); // ID
		String userName = userInfo.getData3(); // name
		String title = userInfo.getData4(); // 职称
		String grade = userInfo.getData5(); // 岗位级别

		boolean go = false;
		if (!userId.trim().equals("") && !userName.trim().equals("")) {
		    go = true;
		}
		String fileName = "D:/bnu/2015/教师考核表_" + userId + "_" + userName + ".doc";
		wordFile = new File(fileName);
		if (wordFile.exists()) {
		    System.out.println("简历存在=======:" + fileName);
		    go = false;
		}

		if (go) {
		    System.out.println("输出简历begin=======:" + userId + ";" + userName);

		    /************ jxgz ***********/
		    StringBuffer jxgz = new StringBuffer();
		    StringBuffer kcxxContent = new StringBuffer();
		    if (kcxxMap != null && kcxxMap.containsKey(userName)) {
			List<DataVo> tempList2 = kcxxMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    kcxxContent.append((j + 1) + ". ");
			    if (!temp2.getData1().equals("")) {
				kcxxContent.append("\"" + temp2.getData1() + "\"");
			    }
			    if (!temp2.getData2().equals("")) {
				kcxxContent.append(", " + temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				kcxxContent.append(", " + temp2.getData3());
			    }
			    if (!temp2.getData4().equals("")) {
				kcxxContent.append(", " + temp2.getData4());
			    }
			    kcxxContent.append("\n");
			}
		    }

		    StringBuffer dsxsContent = new StringBuffer();
		    if (dsxsMap != null && dsxsMap.containsKey(userName)) {
			List<DataVo> tempList2 = dsxsMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    int tempIndex = 1;
			    if (!temp2.getData2().equals("") && !temp2.getData2().equals("0")
				    && !temp2.getData2().equals("#N/A!") && !temp2.getData2().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2011级学术型硕士C：" + temp2.getData2() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData3().equals("") && !temp2.getData3().equals("0")
				    && !temp2.getData3().equals("#N/A!") && !temp2.getData3().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2011级学术型博士：" + temp2.getData3() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData4().equals("") && !temp2.getData4().equals("0")
				    && !temp2.getData4().equals("#N/A!") && !temp2.getData4().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级学术型硕士：" + temp2.getData4() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData5().equals("") && !temp2.getData5().equals("0")
				    && !temp2.getData5().equals("#N/A!") && !temp2.getData5().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级学术型博士：" + temp2.getData5() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData6().equals("") && !temp2.getData6().equals("0")
				    && !temp2.getData6().equals("#N/A!") && !temp2.getData6().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级学术型硕士：" + temp2.getData6() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData7().equals("") && !temp2.getData7().equals("0")
				    && !temp2.getData7().equals("#N/A!") && !temp2.getData7().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级学术型博士：" + temp2.getData7() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData8().equals("") && !temp2.getData8().equals("0")
				    && !temp2.getData8().equals("#N/A!") && !temp2.getData8().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2011级在职硕士：" + temp2.getData8() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData9().equals("") && !temp2.getData9().equals("0")
				    && !temp2.getData9().equals("#N/A!") && !temp2.getData9().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级在职硕士：" + temp2.getData9() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData10().equals("") && !temp2.getData10().equals("0")
				    && !temp2.getData10().equals("#N/A!") && !temp2.getData10().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级在职硕士：" + temp2.getData10() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData11().equals("") && !temp2.getData11().equals("0")
				    && !temp2.getData11().equals("#N/A!") && !temp2.getData11().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级免费师范生教育硕士：" + temp2.getData11() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData12().equals("") && !temp2.getData12().equals("0")
				    && !temp2.getData12().equals("#N/A!") && !temp2.getData12().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级免费师范生教育硕士：" + temp2.getData12() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData13().equals("") && !temp2.getData13().equals("0")
				    && !temp2.getData13().equals("#N/A!") && !temp2.getData13().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2011级英文国际硕士：" + temp2.getData13() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData14().equals("") && !temp2.getData14().equals("0")
				    && !temp2.getData14().equals("#N/A!") && !temp2.getData14().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级英文国际硕士：" + temp2.getData14() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData15().equals("") && !temp2.getData15().equals("0")
				    && !temp2.getData15().equals("#N/A!") && !temp2.getData15().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级英文国际硕士：" + temp2.getData15() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData16().equals("") && !temp2.getData16().equals("0")
				    && !temp2.getData16().equals("#N/A!") && !temp2.getData16().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级英文国际博士：" + temp2.getData16() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData17().equals("") && !temp2.getData17().equals("0")
				    && !temp2.getData17().equals("#N/A!") && !temp2.getData17().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2011级全日制教育硕士：" + temp2.getData17() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData18().equals("") && !temp2.getData18().equals("0")
				    && !temp2.getData18().equals("#N/A!") && !temp2.getData18().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2012级全日制教育硕士：" + temp2.getData18() + "\n");
				tempIndex++;
			    }
			    if (!temp2.getData19().equals("") && !temp2.getData19().equals("0")
				    && !temp2.getData19().equals("#N/A!") && !temp2.getData19().equals("#N/A")) {
				dsxsContent.append(tempIndex + ". 2013级全日制教育硕士：" + temp2.getData19() + "\n");
				tempIndex++;
			    }
			}
		    }

		    if (!kcxxContent.toString().trim().equals("")) {
			jxgz.append("讲授课程：\n");
			jxgz.append(kcxxContent + "\n");
		    }
		    if (!dsxsContent.toString().trim().equals("")) {
			jxgz.append("指导学生：\n");
			jxgz.append(dsxsContent);
		    }
		    /************ jxgz end ***********/

		    /************ kygz ***********/
		    StringBuffer kygz = new StringBuffer();
		    StringBuffer xslwContent = new StringBuffer();
		    if (xslwMap != null && xslwMap.containsKey(userName)) {
			List<DataVo> tempList2 = xslwMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				xslwContent.append((j + 1) + ". " + temp2.getData2());
			    }
			    xslwContent.append("\n");
			}
		    }

		    StringBuffer zzContent = new StringBuffer();
		    if (xszzMap != null && xszzMap.containsKey(userName)) {
			List<DataVo> tempList2 = xszzMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				zzContent.append((j + 1) + ". " + temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				zzContent.append(", <<" + temp2.getData3() + ">>");
			    }
			    if (!temp2.getData4().equals("")) {
				zzContent.append(", " + temp2.getData4());
			    }
			    if (!temp2.getData5().equals("")) {
				zzContent.append(", " + temp2.getData5());
			    }
			    zzContent.append("\n");
			}
		    }

		    StringBuffer kyhjContent = new StringBuffer();
		    if (kyhjMap != null && kyhjMap.containsKey(userName)) {
			List<DataVo> tempList2 = kyhjMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				kyhjContent.append((j + 1) + ". " + temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				kyhjContent.append(", " + temp2.getData3());
			    }
			    if (!temp2.getData4().equals("")) {
				kyhjContent.append(", " + temp2.getData4());
			    }
			    if (!temp2.getData5().equals("")) {
				kyhjContent.append(", " + temp2.getData5());
			    }
			    if (!temp2.getData6().equals("")) {
				kyhjContent.append(", " + temp2.getData6());
			    }
			    kyhjContent.append("\n");
			}
		    }

		    StringBuffer yjbgContent = new StringBuffer();
		    if (yjbgMap != null && yjbgMap.containsKey(userName)) {
			List<DataVo> tempList2 = yjbgMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				yjbgContent.append((j + 1) + ". " + temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				yjbgContent.append(", " + temp2.getData3());
			    }
			    if (!temp2.getData4().equals("")) {
				yjbgContent.append(", 采纳单位：" + temp2.getData4() + ", 2013年");
			    }
			    yjbgContent.append("\n");
			}
		    }

		    StringBuffer kyxmContent = new StringBuffer();
		    if (kyxmMap != null && kyxmMap.containsKey(userName)) {
			List<DataVo> tempList2 = kyxmMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData4().equals("")) {
				kyxmContent.append((j + 1) + ". \"" + temp2.getData4() + "\"");
			    }
			    if (!temp2.getData2().equals("")) {
				kyxmContent.append(", " + temp2.getData2());
			    }
			    if (!temp2.getData3().equals("") && !temp2.getData3().equals("无")) {
				kyxmContent.append(", " + temp2.getData3());
			    }
			    if (!temp2.getData5().equals("")) {
				// 项目批准号
				kyxmContent.append(", " + temp2.getData5());
			    }

			    kyxmContent.append("\n");
			}
		    }

		    if (!xslwContent.toString().trim().equals("")) {
			kygz.append("学术论文：\n");
			kygz.append(xslwContent + "\n");
		    }
		    if (!zzContent.toString().trim().equals("")) {
			kygz.append("著作：\n");
			kygz.append(zzContent + "\n");
		    }
		    if (!kyhjContent.toString().trim().equals("")) {
			kygz.append("奖励：\n");
			kygz.append(kyhjContent + "\n");
		    }
		    if (!yjbgContent.toString().trim().equals("")) {
			kygz.append("研究（咨询）报告：\n");
			kygz.append(yjbgContent + "\n");
		    }
		    if (!kyxmContent.toString().trim().equals("")) {
			kygz.append("2013年科研立项：\n");
			kygz.append(kyxmContent);
		    }
		    /************ kygz end ***********/

		    StringBuffer xzzwContent = new StringBuffer();
		    if (xzzwMap != null && xzzwMap.containsKey(userName)) {
			List<DataVo> tempList2 = xzzwMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				xzzwContent.append(temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				xzzwContent.append(temp2.getData3());
			    }
			}
		    }

		    StringBuffer xxshgzContent = new StringBuffer();
		    if (xxshgzMap != null && xxshgzMap.containsKey(userName)) {
			List<DataVo> tempList2 = xxshgzMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData1().equals("")) {
				xxshgzContent.append(temp2.getData1());
			    }
			}
		    }

		    StringBuffer qtgzContent = new StringBuffer();
		    if (shfwMap != null && shfwMap.containsKey(userName)) {
			List<DataVo> tempList2 = shfwMap.get(userName);
			for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
			    DataVo temp2 = tempList2.get(j);

			    if (!temp2.getData2().equals("")) {
				qtgzContent.append(temp2.getData2());
			    }
			    if (!temp2.getData3().equals("")) {
				qtgzContent.append(temp2.getData3());
			    }
			    qtgzContent.append("\n");
			}
		    }

		    WordTable[] wt = new WordTable[10];
		    for (int arrayIndex = 0; arrayIndex < wt.length; arrayIndex++) {
			wt[arrayIndex] = new WordTable();
			switch (arrayIndex) {
			case 0:// 部门
			    wt[arrayIndex].originalText = "depart";
			    wt[arrayIndex].finalText = dept;
			    break;
			case 1:// 姓名
			    wt[arrayIndex].originalText = "name";
			    wt[arrayIndex].finalText = userName;
			    break;
			case 2:// 工作证号
			    wt[arrayIndex].originalText = "number";
			    wt[arrayIndex].finalText = userId;
			    break;
			case 3:// 职称
			    wt[arrayIndex].originalText = "title";
			    wt[arrayIndex].finalText = title;
			    break;
			case 4:// 岗位级别
			    wt[arrayIndex].originalText = "grade";
			    wt[arrayIndex].finalText = grade;
			    break;
			case 5:// 行政职务
			    wt[arrayIndex].originalText = "xzzw";
			    wt[arrayIndex].finalText = xzzwContent.toString();
			    break;
			case 6:// 担任学校 社会工作
			    wt[arrayIndex].originalText = "work";
			    wt[arrayIndex].finalText = xxshgzContent.toString();
			    break;
			case 7:// 教学情况
			    wt[arrayIndex].originalText = "instructionalState";
			    wt[arrayIndex].finalText = jxgz.toString();
			    break;
			case 8:// 科研工作
			    wt[arrayIndex].originalText = "kygz";
			    wt[arrayIndex].finalText = kygz.toString();
			    break;
			case 9: // 其它工作
			    wt[arrayIndex].originalText = "otherwork";
			    wt[arrayIndex].finalText = qtgzContent.toString();
			    break;
			}
		    }
		    ModifyWordDocument word = new ModifyWordDocument(wt);
		    word.getWord();
		    System.out.println("============输出简历end===================" + i);
		}
	    }
	}

	System.out.println("============Finally....end...happy happy===================");
    }
}
