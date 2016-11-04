package shi2015;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import shi2013.DataVo;
import word.ModifyWordDocument;
import word.WordTable;

public class Shi2015ToWordUsingMap {

	public static void main(String[] args) throws Exception {
		System.out.println("============beigin===============");
		
		List<DataVo> peopeInfoList = new ArrayList<DataVo>();
		DataVo temp = null;
		InputStream is = new FileInputStream("E:/bnu/2015/people20152.xls");
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
			Map<String, List<DataVo>> kcxxMap = null; // 课程信息
			Map<String, List<DataVo>> dsxsMap = null; // 导师带学生
			Map<String, List<DataVo>> shfwMap = null; // 其他工作
			Map<String, List<DataVo>> xzzwMap = null; // 行政职务
			Map<String, List<DataVo>> xxshgzMap = null; // 担任学校社会工作
			Map<String, List<DataVo>> xslwMap = null; // 学术论文
			Map<String, List<DataVo>> xszzMap = null; // 学术著作
			Map<String, List<DataVo>> kyhjMap = null; // 科研获奖
			Map<String, List<DataVo>> yjbgMap = null; // 研究咨询报告
			Map<String, List<DataVo>> kyxmMap = null; // 科研项目

			kcxxMap = Kcxx.getData(false); // 2015全部授课及实习.xls
			dsxsMap = Dsxs.getData(false); // 学生名册.xls
			shfwMap = Shfw.getData(false); // 2015其他工作完成情况.xls
			xzzwMap = Xzzw.getData(false); // 2015年现任行政职务.xls
			xxshgzMap = Xxshgz.getData(false); // 2015年度考核 担任社会工作情况（整理好）.xls
			xslwMap = Xslw.getData(false); // 2015年论文 .xls
			xszzMap = Xszz.getData(false); // 2015年著作.xls
			kyhjMap = Kyhj.getData(false); // 2015年奖励.xls
			yjbgMap = Yjbg.getData(false); // 2015年报告.xls
			kyxmMap = Kyxm.getData(false); // 2015年省部级项目.xls

			boolean outWord = true;

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
				String fileName = "E:/bnu/2015/导出word/教师考核表_" + dept + "_" + userId + "_" + userName + ".doc";
				wordFile = new File(fileName);
				if (wordFile.exists()) {
					System.out.println("简历存在=======:" + fileName);
					go = false;
				}

				/************ 指定输出某人begin **************/
//				go = false;
//				if (userName.equals("石中英") || userName.equals("顾明远")) {
//					go = true;
//				}
				/************ 指定输出某人end **************/

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
							if (!temp2.getData2().trim().equals("") && !temp2.getData2().equals("0") && !temp2.getData2().equals(" ")
									&& !temp2.getData2().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2012级学术型硕士：" + temp2.getData2() + "\n");
								tempIndex++;
							}
							if (!temp2.getData3().trim().equals("") && !temp2.getData3().equals("0") && !temp2.getData3().equals(" ")
									&& !temp2.getData3().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2012级学术型博士：" + temp2.getData3() + "\n");
								tempIndex++;
							}
							if (!temp2.getData4().trim().equals("") && !temp2.getData4().equals("0") && !temp2.getData4().equals(" ")
									&& !temp2.getData4().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级学术型硕士：" + temp2.getData4() + "\n");
								tempIndex++;
							}
							if (!temp2.getData5().trim().equals("") && !temp2.getData5().equals("0") && !temp2.getData5().equals(" ")
									&& !temp2.getData5().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级学术型博士：" + temp2.getData5() + "\n");
								tempIndex++;
							}
							if (!temp2.getData6().trim().equals("") && !temp2.getData6().equals("0") && !temp2.getData6().equals(" ")
									&& !temp2.getData6().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级学术型硕士：" + temp2.getData6() + "\n");
								tempIndex++;
							}
							if (!temp2.getData7().trim().equals("") && !temp2.getData7().equals("0") && !temp2.getData7().equals(" ")
									&& !temp2.getData7().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级学术型博士：" + temp2.getData7() + "\n");
								tempIndex++;
							}
							if (!temp2.getData8().trim().equals("") && !temp2.getData8().equals("0") && !temp2.getData8().equals(" ")
									&& !temp2.getData8().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级暑期教育硕士：" + temp2.getData8() + "\n");
								tempIndex++;
							}
							if (!temp2.getData9().trim().equals("") && !temp2.getData9().equals("0") && !temp2.getData9().equals(" ")
									&& !temp2.getData9().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级暑期教育硕士：" + temp2.getData9() + "\n");
								tempIndex++;
							}
							if (!temp2.getData10().trim().equals("") && !temp2.getData10().equals("0") && !temp2.getData10().equals(" ")
									&& !temp2.getData10().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2012级英文国际硕士：" + temp2.getData10() + "\n");
								tempIndex++;
							}
							if (!temp2.getData11().trim().equals("") && !temp2.getData11().equals("0") && !temp2.getData11().equals(" ")
									&& !temp2.getData11().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级英文国际硕士：" + temp2.getData11() + "\n");
								tempIndex++;
							}
							if (!temp2.getData12().trim().equals("") && !temp2.getData12().equals("0") && !temp2.getData12().equals(" ")
									&& !temp2.getData12().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级英文国际博士：" + temp2.getData12() + "\n");
								tempIndex++;
							}
							if (!temp2.getData13().trim().equals("") && !temp2.getData13().equals("0") && !temp2.getData13().equals(" ")
									&& !temp2.getData13().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级英文国际硕士：" + temp2.getData13() + "\n");
								tempIndex++;
							}
							if (!temp2.getData14().trim().equals("") && !temp2.getData14().equals("0") && !temp2.getData14().equals(" ")
									&& !temp2.getData14().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级英文国际博士：" + temp2.getData14() + "\n");
								tempIndex++;
							}
							if (!temp2.getData15().trim().equals("") && !temp2.getData15().equals("0") && !temp2.getData15().equals(" ")
									&& !temp2.getData15().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2013级全日制教育硕士：" + temp2.getData15() + "\n");
								tempIndex++;
							}
							if (!temp2.getData16().trim().equals("") && !temp2.getData16().equals("0") && !temp2.getData16().equals(" ")
									&& !temp2.getData16().equals("#N/A")) {
								dsxsContent.append(tempIndex + ". 2014级全日制教育硕士：" + temp2.getData16() + "\n");
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
							// 1A,2B,3C,4D,5E,6F,7G,8H,9I,10J,11K,12L,13M,14N,15O,16P,17Q,18R,19S,20T,21U,22V,23W,24X.

							if (!temp2.getData3().equals("")) {
								xslwContent.append((j + 1) + ". " + temp2.getData3().trim() + ". ");
							}
							if (!temp2.getData2().equals("")) {
								xslwContent.append(temp2.getData2().trim());
							}
							if (!temp2.getData9().equals("")) {
								xslwContent.append(temp2.getData9().trim() + ". ");
							}
							if (!temp2.getData8().equals("")) {
								xslwContent.append(temp2.getData8().trim() + "， ");
							}
							if (!temp2.getData10().equals("")) {
								xslwContent.append(temp2.getData10().trim());
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
								yjbgContent.append(", 采纳单位：" + temp2.getData4() + ", 2014年");
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
						kygz.append("2014年科研立项：\n");
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
						System.out.println("============输出简历end===================" + i);
					} else {
						System.out.println("部门:" + dept + ";姓名:" + userName + ";工作证号:" + userId + ";职称:" + title + ";岗位级别:" + grade);
						System.out.println("行政职务:" + xzzwContent.toString());
						System.out.println("担任学校 社会工作:" + xxshgzContent.toString());
						System.out.println("教学情况:\n" + jxgz.toString());
						System.out.println("科研工作:\n" + kygz.toString());
						System.out.println("其它工作:" + qtgzContent.toString());
						System.out.println("--------------------------------------------------------------------------------");
					}
				}
			}
		}

		System.out.println("============Finally....end...happy happy===================");
	}
}
