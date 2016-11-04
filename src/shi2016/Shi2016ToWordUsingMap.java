package shi2016;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;
import word.ModifyWordDocument;
import word.WordTable;

public class Shi2016ToWordUsingMap {

	public static void main(String[] args) throws Exception {
		System.out.println("============beigin===============");

		String exportWord = "E:/bnu/2016/导出word/教师考核表_";
		boolean outWord = true; // 是否输出word。测试时为false
		String peoplePath = "E:/bnu/2016/people2016.xls";

		List<DataVo> peopeInfoList = new ArrayList<DataVo>();
		DataVo temp = null;
		InputStream is = new FileInputStream(peoplePath);
		Workbook rwb = Workbook.getWorkbook(is);
		Sheet rs = rwb.getSheet(0);
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

				// System.out.println(temp.toString());

				peopeInfoList.add(temp);
			}
		}
		rwb.close();
		is.close();
		if (peopeInfoList == null || peopeInfoList.size() <= 0) {
		} else {
			System.out.println("============人员信息====获取成功！！！====总人数为=" + peopeInfoList.size());
		}

		if (peopeInfoList != null && peopeInfoList.size() > 0) {
			String dsxsPath = "E:/bnu/2016/学生名册汇总（给田桦）.xls"; // 导师带学生
			String xslwPath = "E:/bnu/2016/2016年科研/论文.xls"; // 学术论文
			String xszzPath = "E:/bnu/2016/2016年科研/著作.xls"; // 学术著作
			String kyhjPath = "E:/bnu/2016/2016年科研/获奖.xls"; // 科研获奖
			String yjbgPath = "E:/bnu/2016/2016年科研/咨询报告.xls"; // 研究咨询报告
			String kyxmPath = "E:/bnu/2016/2016年科研/省部级项目.xls"; // 科研项目
			String shfwPath = "E:/bnu/2016/2016社会服务/其他工作完成情况.xls"; // 其他工作
			String xzzwPath = "E:/bnu/2016/2016社会服务/现任行政职务.xls"; // 行政职务
			String xxshgzPath = "E:/bnu/2016/2016社会服务/担任社会工作情况（整理好）.xls"; // 担任学校社会工作

			String kcxxPath2 = "E:/bnu/2016/刘烨汇总（给田桦）(1).xls"; // 课程信息
			String kcxxPath3 = "E:/bnu/2016/夜大（给田桦）.xls"; // 课程信息
			String kcxxPath4 = "E:/bnu/2016/实习合并（给田桦）.xls"; // 课程信息

			Map<String, List<DataVo>> kcxxMap2 = Kcxx2.getData(false, kcxxPath2); // 课程信息
			Map<String, List<DataVo>> kcxxMap3 = Kcxx3.getData(false, kcxxPath3); // 课程信息
			Map<String, List<DataVo>> kcxxMap4 = Kcxx4.getData(false, kcxxPath4); // 课程信息
			// Map<String, List<DataVo>> dsxsMap = Dsxs.getData(false,
			// dsxsPath); // 导师带学生
			Map<String, Map<String, List<DataVo>>> dsxsMap = Dsxs2.getData(false, dsxsPath); // 导师带学生
			Map<String, List<DataVo>> shfwMap = Shfw.getData(false, shfwPath); // 其他工作
			Map<String, List<DataVo>> xzzwMap = Xzzw.getData(false, xzzwPath); // 行政职务
			Map<String, List<DataVo>> xxshgzMap = Xxshgz.getData(false, xxshgzPath); // 担任学校社会工作
			Map<String, List<DataVo>> xslwMap = Xslw.getData(false, xslwPath); // 学术论文
			Map<String, List<DataVo>> xszzMap = Xszz.getData(false, xszzPath); // 学术著作
			Map<String, List<DataVo>> kyhjMap = Kyhj.getData(false, kyhjPath); // 科研获奖
			Map<String, List<DataVo>> yjbgMap = Yjbg.getData(false, yjbgPath); // 研究咨询报告
			Map<String, List<DataVo>> kyxmMap = Kyxm.getData(false, kyxmPath); // 科研项目

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

			File wordFile = null;
			for (int i = 0; i < peopeInfoList.size(); i++) {
				DataVo userInfo = peopeInfoList.get(i);
				String dept = userInfo.getData1(); // 部门
				String userId = userInfo.getData2(); // ID
				String userName = userInfo.getData3(); // name
				String title = userInfo.getData4(); // 职称
				String grade = userInfo.getData5(); // 岗位级别

				boolean go = false;
				if (!userId.trim().equals("") && !userName.trim().equals("")) {
					go = true;
				}
				String fileName = exportWord + dept + "_" + userId + "_" + userName + ".doc";
				wordFile = new File(fileName);
				if (wordFile.exists()) {
					System.out.println("简历存在=======:" + fileName);
					go = false;
				}

				/************************************************** 指定输出某人begin ****************************************************/
//				go = false;
//				if (userName.equals("石中英") || userName.equals("顾明远") || userName.equals("薛二勇")
//						|| userName.equals("朱旭东") || userName.equals("黄荣怀")
//						|| userName.equals("陈桄") || userName.equals("余胜泉")) {
//					go = true;
//				}
				/************************************************** 指定输出某人end ****************************************************/

				if (go) {
					System.out.println("输出简历begin=======:" + userId + ";" + userName);

					/************************************************** jxgz *************************************************/
					StringBuffer jxgz = new StringBuffer();
					StringBuffer kcxxContent = new StringBuffer();
					int jxgzIndex = 1;
					if (kcxxMap2 != null && kcxxMap2.containsKey(userName)) {
						List<DataVo> tempList2 = kcxxMap2.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);

							kcxxContent.append(jxgzIndex + ". ");
							if (!temp2.getData1().equals("")) {
								kcxxContent.append("\"" + temp2.getData1() + "\"");
							}
							if (!temp2.getData2().equals("")) {
								kcxxContent.append(", " + temp2.getData2());
							}
							kcxxContent.append("\n");
							jxgzIndex++;
						}
					}
					if (kcxxMap3 != null && kcxxMap3.containsKey(userName)) {
						List<DataVo> tempList2 = kcxxMap3.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);

							kcxxContent.append(jxgzIndex + ". ");
							if (!temp2.getData1().equals("")) {
								kcxxContent.append("\"" + temp2.getData1() + "\"");
							}
							if (!temp2.getData2().equals("")) {
								kcxxContent.append(", " + temp2.getData2());
							}
							if (!temp2.getData3().equals("")) {
								kcxxContent.append(", " + temp2.getData3());
							}
							kcxxContent.append("\n");
							jxgzIndex++;
						}
					}
					if (kcxxMap4 != null && kcxxMap4.containsKey(userName)) {
						List<DataVo> tempList2 = kcxxMap4.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);

							kcxxContent.append(jxgzIndex + ". ");
							if (!temp2.getData1().equals("")) {
								kcxxContent.append("\"" + temp2.getData1() + "\"");
							}
							if (!temp2.getData2().equals("")) {
								kcxxContent.append(", " + temp2.getData2());
							}
							if (!temp2.getData3().equals("")) {
								kcxxContent.append(", " + temp2.getData3());
							}
							kcxxContent.append("\n");
							jxgzIndex++;
						}
					}

					StringBuffer dsxsContent = new StringBuffer();
					if (dsxsMap != null && dsxsMap.containsKey(userName)) {
						Map<String, List<DataVo>> dataMap = dsxsMap.get(userName);
						if (dataMap != null) {
							int index = 1;
							for (String base : baseList) {
								if (dataMap.containsKey(base)) {
									List<DataVo> dataList = dataMap.get(base);
									if (dataList != null && dataList.size() > 0) {
										dsxsContent.append(index + ". " + base + ": ");
										for (int dsxsLoop = 0; dsxsLoop < dataList.size(); dsxsLoop++) {
											DataVo dsxsDataTemp = dataList.get(dsxsLoop);
											if (dsxsLoop == 0) {
												dsxsContent.append(dsxsDataTemp.getData1());
											} else {
												dsxsContent.append(", " + dsxsDataTemp.getData1());
											}
										}
										dsxsContent.append("\n");
										index++;
									}
								}
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
					/************************************************** jxgz end *************************************************/

					/************************************************** kygz *************************************************/
					StringBuffer kygz = new StringBuffer();
					StringBuffer xslwContent = new StringBuffer();
					if (xslwMap != null && xslwMap.containsKey(userName)) {
						List<DataVo> tempList2 = xslwMap.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);
							// 2,3,10,11,12
							if (!temp2.getData3().equals("")) {
								xslwContent.append((j + 1) + ". " + temp2.getData3().trim() + ". ");
							}
							if (!temp2.getData2().equals("")) {
								xslwContent.append(temp2.getData2().trim());
							}
							if (!temp2.getData11().equals("")) {
								xslwContent.append(temp2.getData11().trim() + ". ");
							}
							if (!temp2.getData10().equals("")) {
								xslwContent.append(temp2.getData10().trim() + "，");
							}
							if (!temp2.getData12().equals("")) {
								xslwContent.append(temp2.getData12().trim());
							}
							xslwContent.append("\n");
						}
					}

					StringBuffer zzContent = new StringBuffer();
					if (xszzMap != null && xszzMap.containsKey(userName)) {
						List<DataVo> tempList2 = xszzMap.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);

							if (!temp2.getData4().equals("")) {
								zzContent.append((j + 1) + ". " + temp2.getData4());
							}
							if (!temp2.getData5().equals("")) {
								zzContent.append(". <<" + temp2.getData5() + ">>");
							}
							if (!temp2.getData6().equals("")) {
								zzContent.append(", " + temp2.getData6());
							}
							if (!temp2.getData7().equals("")) {
								zzContent.append(", " + temp2.getData7());
							}
							zzContent.append("\n");
						}
					}

					StringBuffer kyhjContent = new StringBuffer();
					if (kyhjMap != null && kyhjMap.containsKey(userName)) {
						List<DataVo> tempList2 = kyhjMap.get(userName);
						for (int j = 0; tempList2 != null && j < tempList2.size(); j++) {
							DataVo temp2 = tempList2.get(j);

							if (!temp2.getData1().equals("")) {
								kyhjContent.append((j + 1) + ". " + temp2.getData1());
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
							if (!temp2.getData7().equals("")) {
								kyhjContent.append(", " + temp2.getData7());
							}
							kyhjContent.append(", 2015年");
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
								yjbgContent.append(", 采纳单位：" + temp2.getData4() + ", 2015年");
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
						kygz.append("2015年科研立项：\n");
						kygz.append(kyxmContent);
					}
					/************************************************** kygz end *************************************************/

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

					if (outWord) {
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
						System.out.println("============输出简历end==========" + i);
					} else {
						System.out.println("部门:" + dept + ";姓名:" + userName + ";工作证号:" + userId
								+ ";职称:" + title + ";岗位级别:" + grade);
						System.out.println("行政职务:" + xzzwContent.toString());
						System.out.println("担任学校 社会工作:" + xxshgzContent.toString());
						System.out.println("教学情况:\n" + jxgz.toString());
						System.out.println("科研工作:\n" + kygz.toString());
						System.out.println("其它工作:" + qtgzContent.toString());
						System.out.println("------------------------------------------------");
					}
				}
			}
		}

		System.out.println("============Finally....end...happy happy===================");
	}
}
