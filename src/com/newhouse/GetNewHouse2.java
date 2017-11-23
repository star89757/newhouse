package com.newhouse;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.http.client.config.CookieSpecs;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class GetNewHouse2 {

	static int pageSize = -1;//提取多少页，如果-1，则提取所有
	static int page = 1;
	static int timeOut = 20000;
	static int count = 0;
	static int finishCount_detail = 0;
	static int finishCount_dt = 0;
	static int finishCount_hx = 0;
	static int finishCount_kpxq = 0;
	static int dd = 0;
	
	static List<Map<String,String>> houseList = new ArrayList<Map<String,String>>();
	
	public static void main(String[] args) {
		
		//String htmlUrl = "http://newhouse.dg.fang.com/house/s";
		String htmlUrl = "";
		
		System.out.println("请输入链接地址：");
		Scanner scan = new Scanner(System.in);
		htmlUrl = scan.nextLine();
		
		if(!"".equals(htmlUrl)){
			getNewHouse(htmlUrl,htmlUrl,houseList);
			
			//所有线程执行完后写入Excel文件
			while(count != finishCount_detail || count != finishCount_dt || count != finishCount_hx || count != finishCount_kpxq){
				try {
					//System.out.println("等待所有网页提取完成！");
					Thread.sleep(2000);
				} catch (InterruptedException e) {
				}
			}
			String fileName = exportExcel(houseList);
			System.out.println("===============成功生成文件：" + fileName + "===============");
			System.out.println("===============执行完成===============");
		}
	}

	/**
	 * 导出Excel文件
	 * @param excelFile
	 * @param houseList
	 * @return 
	 */
	@SuppressWarnings("deprecation")
	private static String exportExcel(List<Map<String, String>> houseList) {
		
		FileOutputStream out = null;
		String fileName = "";
		File destFile = null;
		
		try {
			//定义文件名
			fileName = "newhouse_" + new SimpleDateFormat("yyyyMMddHHmmssS").format(new Date())+".xlsx";
			//String src = "blank.xlsx";
			
			destFile = new File(String.format("%s%s%s", "d:\\" , File.separator , fileName));
			//FileUtils.copyFile(new File(src), destFile);
			
			//生成Excel报表
			Workbook wb = new SXSSFWorkbook(10000);
			Map<String, CellStyle> styleMap = ExcelUtil.getCommonStyle(wb);
			CellStyle linkStyle  = wb.createCellStyle();
			linkStyle.cloneStyleFrom(styleMap.get("stringStyle"));
			
			Font f = wb.createFont();
			f.setUnderline(Font.U_SINGLE);
			f.setColor(HSSFColor.BLUE.index);
			linkStyle.setFont(f);
			
			CellStyle wrapStyle  = wb.createCellStyle();
			wrapStyle.cloneStyleFrom(styleMap.get("stringStyle"));
			wrapStyle.setWrapText(true);
			
			//------------------------------------------------写明细报表数据------------------------------------------------//
			Sheet mxSheet = wb.createSheet();
			wb.setSheetName(0, "newhouse");

			List<String> mxHeads = new ArrayList<String>();//明细报表表头数据
			mxHeads.add("名称");
			mxHeads.add("别名");
			mxHeads.add("价格");
			
			mxHeads.add("物业类别");
			mxHeads.add("产权年限");
			mxHeads.add("开发商");
			mxHeads.add("楼盘地址");
			mxHeads.add("占地面积");
			
			mxHeads.add("容积率");
			mxHeads.add("建筑面积");
			mxHeads.add("楼栋总数");
			mxHeads.add("总户数");
			mxHeads.add("停车位");
			
			mxHeads.add("户型");
			mxHeads.add("物业费");
			mxHeads.add("物业公司");
			mxHeads.add("楼层状况");
			
			//mxHeads.add("楼盘动态标题");
			mxHeads.add("楼盘动态内容");
			mxHeads.add("开盘时间");
			mxHeads.add("开盘详情");
			mxHeads.add("交房时间");
			
			mxHeads.add("项目简介");
			
			mxHeads.add("网址");
			
			//写表头
			Row mxRow = mxSheet.createRow(0);
			for(int i = 0;i< mxHeads.size();i++){
				mxSheet.setColumnWidth(i, 4500);
				ExcelUtil.writeCell(mxRow,i,Cell.CELL_TYPE_STRING,mxHeads.get(i),styleMap.get("headStyle"));
			}
			
			//写数据体数据
			for(int i = 0;i<houseList.size();i++){
				Map<String,String> data = houseList.get(i);
				Row dataRow = mxSheet.createRow(mxSheet.getLastRowNum() + 1);
				int j = 0;
				for(int z = 0;z<mxHeads.size();z++){
					if("网址".equals(mxHeads.get(z))){
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),linkStyle);
						Cell cell = dataRow.getCell(j-1);
						cell.setCellType(Cell.CELL_TYPE_FORMULA);
						cell.setCellFormula("HYPERLINK(\"" + data.get(mxHeads.get(z)) + "\",\"" + data.get(mxHeads.get(z)) + "\")");
					}/*else if("户型".equals(mxHeads.get(z))){
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),wrapStyle);
					}*/else{
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),styleMap.get("stringStyle"));
					}
				}
			}
			
			//写Excel文件
			out = new FileOutputStream(destFile);
			out.flush();
			wb.write(out);
			out.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally{
			if(out != null)
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
		return destFile.getPath();
	}

	/**
	 * 获取楼盘信息
	 * @param htmlUrl
	 * @param htmlUrl2
	 */
	public static void getNewHouse(String baseUrl, String htmlUrl,List<Map<String,String>> houseList) {
		System.out.println("=========正在执行第[" + page + "]页=========");
		try {
			//获取当前页的楼盘信息
			//Document doc  = Jsoup.connect(htmlUrl).timeout(timeOut).get();
			Document doc = getDoc(htmlUrl);
			if(null == doc){
				return;
			}
			Elements nhouse_list = doc.getElementsByClass("nhouse_list");
			if(nhouse_list.size() > 0){
				Elements nl_con = nhouse_list.get(0).getElementsByClass("nl_con");
				for(int i = 0;i<nl_con.size();i++){
					Elements nhouse_li = nl_con.get(i).getElementsByTag("li");
					for(int j = 0;j<nhouse_li.size();j++){
						Elements nhouse_li_a = nhouse_li.get(j).getElementsByTag("a");
						if(nhouse_li_a.size() > 1){
							String houseName = nhouse_li_a.get(1).text();
							String houseLink = nhouse_li_a.get(1).attr("href");
							//Document detailDoc_1  = Jsoup.connect(houseLink).timeout(timeOut).get();
							Map<String,String> houseMap = new HashMap<String,String>();
							directionToHouseDetail(houseLink,houseName,houseMap);
						}
					}
				}
			}else{
				//一些老的特殊网页,如“梅州”等二线城市
				Elements sslist = doc.getElementsByClass("sslist");
				if(sslist.size()> 0){
					Elements sslalone = sslist.get(0).getElementsByClass("sslalone");
					for(int i =0;i<sslalone.size();i++){
						String houseName = "";
						String houseLink = "";
						
						try {
							houseName = sslalone.get(i).getElementsByClass("sslainfor").get(0).getElementsByTag("a").get(0).text();
							houseLink = sslalone.get(i).getElementsByClass("sslainfor").get(0).getElementsByTag("a").get(0).attr("href");
						} catch (Exception e) {
						}
						if(!"".equals(houseName)){
							Map<String,String> houseMap = new HashMap<String,String>();
							directionToHouseDetail(houseLink,houseName,houseMap);
						}
					}
				}
			}
			
			//判断是否有下一页(筛选栏中的上下页，有些网页没有)
			if(nhouse_list.size() > 0){
				//新版网页
				Elements otherpage = doc.getElementsByClass("otherpage");
				if(otherpage.size()>0){
					Elements otherpage_a = otherpage.get(0).getElementsByTag("a");
					for(int i = 0;i<otherpage_a.size();i++){
						if(">".equals(otherpage_a.get(i).text())){
							if(-1 == pageSize){
								Thread.sleep(10000);
								page++;
								getNewHouse(baseUrl, baseUrl + otherpage_a.get(i).attr("href"),houseList);
							}else{
								if(page < pageSize){
									Thread.sleep(10000);
									page++;
									getNewHouse(baseUrl, baseUrl + otherpage_a.get(i).attr("href"),houseList);
								}
							}
						}
					}
				}else{
					//页脚上下页
					Elements nextPage = doc.getElementsByClass("page");
					if(nextPage.size()>0){
						Elements nextPage_a = nextPage.get(0).getElementsByClass("next");
						if(nextPage_a.size() > 0){
							String nextPageLink = nextPage_a.get(0).attr("href");
							if(-1 == pageSize){
								Thread.sleep(10000);
								page++;
								getNewHouse(baseUrl, baseUrl + nextPageLink,houseList);
							}else{
								if(page < pageSize){
									Thread.sleep(10000);
									page++;
									getNewHouse(baseUrl, baseUrl + nextPageLink,houseList);
								}
							}
						}
					}
				}
			}else{
				//老版网页
				Elements pagearrowright = doc.getElementsByClass("pagearrowright");
				if(pagearrowright.size() > 0){
					String pagearrowrightLink = "";
					try {
						pagearrowrightLink = pagearrowright.get(0).getElementsByTag("a").get(0).attr("href");
					} catch (Exception e) {
					}
					if(!"".equals(pagearrowrightLink)){
						if(-1 == pageSize){
							Thread.sleep(10000);
							page++;
							getNewHouse(baseUrl, baseUrl + pagearrowrightLink,houseList);
						}else{
							if(page < pageSize){
								Thread.sleep(10000);
								page++;
								getNewHouse(baseUrl, baseUrl + pagearrowrightLink,houseList);
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/*
	 * 跳转到“楼盘详情”页面
	 */
	public static void directionToHouseDetail(String houseLink, String houseName, Map<String, String> houseMap){
		Document detailDoc_1  = getDoc(houseLink);
		if(null != detailDoc_1){
			//楼盘详细页面
			Elements navleft = detailDoc_1.getElementsByClass("navleft");
			String detailLink = "";//楼盘详情
			String dtLink  = "";//楼盘动态
			String hxLink = "";//户型
			String kpxqLink = "";//开盘详情
			if(null != navleft && navleft.size() > 0){
				Elements navleft_a = navleft.get(0).getElementsByTag("a");
				for(int x = 0;x<navleft_a.size();x++){
					String linkName = navleft_a.get(x).text();
					if(linkName.indexOf("楼盘详情") != -1 || linkName.indexOf("详细信息") != -1){
						detailLink = navleft_a.get(x).attr("href");
						dtLink = detailLink.replace("housedetail.htm", "dongtai.htm");
						kpxqLink = detailLink.replace("housedetail.htm", "sale_history.htm");
					}
					if(linkName.indexOf("户型") != -1){
						hxLink = navleft_a.get(x).attr("href");
					}
				}
			}
			houseMap.put("名称", houseName);
			houseMap.put("网址", detailLink);
			
			if(!"".equals(detailLink)){
				count++;
				ExecuteThread e1 = new ExecuteThread(detailLink, houseName,"楼盘详情",houseMap);
				e1.start();
				ExecuteThread e2 = new ExecuteThread(dtLink, houseName,"楼盘动态",houseMap);
				e2.start();
				ExecuteThread e3 = new ExecuteThread(hxLink, houseName,"户型",houseMap);
				e3.start();
				ExecuteThread e4 = new ExecuteThread(kpxqLink, houseName,"开盘详情",houseMap);
				e4.start();
				
				if(houseMap.size() > 0){
					houseList.add(houseMap);
				}
			}
		}
	}
	
	/**
	 * 根据url，获取jsoup的document
	 * @param htmlUrl
	 * @return
	 */
	public static Document getDoc(String htmlUrl) {
		String html = getHtml(htmlUrl);
		Document doc = null;
		if(!"".equals(html)){
			doc = Jsoup.parse(html);
		}
		return doc;
	}
	
	/**
	 * 使用htmlclient获取html
	 * @return
	 */
	public static String getHtml(String htmlUrl) {
		// HttpClient 超时配置
		RequestConfig globalConfig = RequestConfig.custom()
				.setCookieSpec(CookieSpecs.STANDARD)
				.setConnectionRequestTimeout(timeOut)
				.setConnectTimeout(timeOut).build();
		CloseableHttpClient httpClient = HttpClients.custom()
				.setDefaultRequestConfig(globalConfig).build();

		// 创建一个GET请求
		HttpGet httpGet = new HttpGet(htmlUrl);
		httpGet.addHeader(
				"User-Agent",
				"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.152 Safari/537.36");
		httpGet.addHeader(
				"Cookie",
				"_gat=1; nsfw-click-load=off; gif-click-load=on; _ga=GA1.2.1861846600.1423061484");
		CloseableHttpResponse response = null;
		try {
			// 发送请求，并执行
			response = httpClient.execute(httpGet);
			InputStream in = response.getEntity().getContent();
			String html = convertStreamToString(in);
			Document doc = Jsoup.parse(html);
			return doc.html();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				response.close();
				httpClient.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return "";
	}
	
	/**
	 * 转换String
	 * @param is
	 * @return
	 */
	public static String convertStreamToString(InputStream is) {
		BufferedReader reader = new BufferedReader(new InputStreamReader(is));
		StringBuilder sb = new StringBuilder();

		String line = null;
		try {
			while ((line = reader.readLine()) != null) {
				sb.append(line);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return sb.toString();
	}
	
	static class ExecuteThread extends Thread{
		private String houseLink;
		private String houseName;
		private String type;
		private Map<String,String> houseMap;
		
		public ExecuteThread(String url,String name,String type,Map<String,String> houseMap){
			this.houseLink = url;
			this.houseName = name;
			this.type = type;
			this.houseMap = houseMap;
		}
		
		public void run(){
			System.out.println("=========正在提取[" + houseName + "][" + type + "][" + houseLink + "]=========");
			try {
				if("楼盘详情".equals(this.type)){
					getHouseDetail(houseLink, houseMap);
				}
				if("楼盘动态".equals(this.type)){
					getHouseDt(houseLink, houseMap);
				}
				if("户型".equals(this.type)){
					getHx(houseLink, houseMap);
				}
				if("开盘详情".equals(this.type)){
					getKpxq(houseLink, houseMap);
				}
			} catch (Exception e) {
				//e.printStackTrace();
				System.out.println("=========提取[" + houseName + "][" + type + "][失败]：" + e.getMessage() + "=========");
			}
		}

		/**
		 * 获取开盘详情
		 * @param kpxqLink
		 * @param houseMap
		 * @throws Exception
		 */
		private void getKpxq(String kpxqLink,
				Map<String, String> houseMap) throws Exception {
				try {
					if(null != kpxqLink && !"".equals(kpxqLink)){
						Thread.sleep(1000);
						Document hxDoc = getDoc(kpxqLink);
						if(null != hxDoc){
							Elements kpjjlu = hxDoc.getElementsByClass("kpjjlu");
							if(kpjjlu.size() > 0){
								Elements tr = kpjjlu.get(0).getElementsByTag("tr");
								String kpjjluStr = "";
								for(int i = 1;i<tr.size();i++){
									String title = "";
									String content = "";
									try {
										title = tr.get(i).getElementsByTag("td").get(0).text();
										content = tr.get(i).getElementsByTag("td").get(1).text();
									} catch (Exception e) {
									}
									
									if(!"".equals(title)){
										kpjjluStr += title + " : " + content + "--";
									}
								}
								houseMap.put("开盘详情", kpjjluStr);
							}
						}
					}
				} catch (Exception e) {
					throw e;
				} finally {
					finishCount_kpxq++;
				}
			}
		
		/**
		 * 获取户型
		 * @param navleft_a
		 * @param houseMap
		 * @throws Exception
		 */
		private void getHx(String hxLink,
			Map<String, String> houseMap) throws Exception {
			try {
				if(null != hxLink && !"".equals(hxLink)){
					Thread.sleep(1000);
					Document hxDoc = getDoc(hxLink);
					if(null != hxDoc){
						Elements xc_img_list = hxDoc.getElementsByClass("xc_img_list");
						String hxStr = "";
						for(int i = 0;i<xc_img_list.size();i++){
							Elements tiaojian = xc_img_list.get(i).getElementsByClass("tiaojian");
							if(tiaojian.size() > 0){
								hxStr += tiaojian.get(0).text() + "--";
							}
						}
						houseMap.put("户型", hxStr);
					}
				}
			} catch (Exception e) {
				throw e;
			} finally {
				finishCount_hx++;
			}
		}
		
		/**
		 * 获取动态
		 * @param detailLink
		 * @param houseMap
		 * @throws Exception
		 */
		private void getHouseDt(String dtLink,
				Map<String, String> houseMap) throws Exception {
			try {
				Thread.sleep(1000);
				Document dtDoc = getDoc(dtLink);
				if(null != dtDoc){
					Elements dtli = dtDoc.getElementsByTag("li");
					for(int i = 0;i<dtli.size();i++){
						Elements time_wrapper = dtli.get(i).getElementsByClass("time-wrapper");
						if(time_wrapper.size() > 0){
							//旧版楼盘动态
							Elements r_content = dtli.get(i).getElementsByClass("r-content");
							if(time_wrapper.size() > 0){
								String title = "";
								try {
									title = time_wrapper.get(0).getElementsByTag("h3").get(0).text() + " " + 
											time_wrapper.get(0).getElementsByTag("h2").get(0).text() + " " + 
											time_wrapper.get(0).getElementsByTag("h1").get(0).text() + " " + 
											r_content.get(0).getElementsByTag("h1").text();
								} catch (Exception e) {
								}
								if(!"".equals(title)){
									houseMap.put("楼盘动态标题", title);
									String dtHref = r_content.get(0).getElementsByTag("a").get(0).attr("href");
									Thread.sleep(1000);
									Document dtDoc_2 = getDoc(dtHref);
									if(null != dtDoc_2){
										Elements atc_wrapper = dtDoc_2.getElementsByClass("atc-wrapper");
										if(atc_wrapper.size() > 0){
											houseMap.put("楼盘动态内容", atc_wrapper.get(0).text());
										}
									}
								}
								break;
							}
						}else{
							//使用新版楼盘动态
							if("storyList".equals(dtli.get(i).attr("class"))){
								Elements storyList_a = dtli.get(i).getElementsByTag("a");
								if(storyList_a.size() > 0){
									String title = "";
									try {
										title = dtli.get(i).getElementsByClass("sLTime").get(0).text() + ":" + storyList_a.get(0).text();
									} catch (Exception e) {
									}
									if(!"".equals(title)){
										houseMap.put("楼盘动态标题", title);
										String dtHref = storyList_a.get(0).attr("href");
										Thread.sleep(1000);
										Document dtDoc_2 = getDoc(dtHref);
										if(null != dtDoc_2){
											Elements atc_wrapper = dtDoc_2.getElementsByClass("atc-wrapper");
											if(atc_wrapper.size() > 0){
												houseMap.put("楼盘动态内容", atc_wrapper.get(0).text());
											}
										}
									}
									break;
								}
							}
						}
						
					}
				}
			} catch (Exception e) {
				throw e;
			} finally {
				finishCount_dt++;
			}
		}
		/**
		 * 获取楼盘详情和楼盘动态
		 * @param navleft_a
		 * @param houseMap
		 * @throws Exception 
		 */
		private void getHouseDetail(String detailLink,
				Map<String, String> houseMap) throws Exception {
			try {
				houseMap.put("名称", houseName);
				houseMap.put("网址", detailLink);
				
				//读取“详细信息”中的内容
				//Document detailDoc_2  = Jsoup.connect(detailLink).timeout(timeOut).get();
				Thread.sleep(1000);
				Document detailDoc_2 = getDoc(detailLink);
				if(null != detailDoc_2){
					Elements h1_label = detailDoc_2.getElementsByClass("h1_label");
					for(int i = 0 ;i<h1_label.size();i++){
						String biemin = h1_label.get(i).text();
						if(null != biemin && biemin.indexOf("别名") != -1){
							houseMap.put("别名", biemin);
							break;
						}
					}
					
					Elements main_left = detailDoc_2.getElementsByClass("main-left");
					if(null != main_left && main_left.size() > 0){
						//价格
						Elements main_info_price = main_left.get(0).getElementsByClass("main-info-price");
						if(null != main_info_price && main_info_price.size() > 0){
							houseMap.put("价格", main_info_price.get(0).text());
						}
						//项目简介
						Elements intro = main_left.get(0).getElementsByClass("intro");
						if(intro.size() > 0){
							houseMap.put("项目简介", intro.get(0).text());
						}
						
						Elements main_left_li = main_left.get(0).getElementsByTag("li");
						for(int z = 0;z < main_left_li.size();z++){
							String left = (null != main_left_li.get(z).getElementsByClass("list-left") && main_left_li.get(z).getElementsByClass("list-left").size() > 0) ? main_left_li.get(z).getElementsByClass("list-left").text() : "";
							left = left.replaceAll(" ", "");
							String right = (null != main_left_li.get(z).getElementsByClass("list-right") && main_left_li.get(z).getElementsByClass("list-right").size() > 0) ? main_left_li.get(z).getElementsByClass("list-right").text() : "";
							if("".equals(right)){
								right = (null != main_left_li.get(z).getElementsByClass("list-right-floor") && main_left_li.get(z).getElementsByClass("list-right-floor").size() > 0) ? main_left_li.get(z).getElementsByClass("list-right-floor").text() : "";
							}
							if("".equals(right)){
								right = (null != main_left_li.get(z).getElementsByClass("list-right-text") && main_left_li.get(z).getElementsByClass("list-right-text").size() > 0) ? main_left_li.get(z).getElementsByClass("list-right-text").text() : "";
							}
							if(!"".equals(left)){
								houseMap.put(left.replaceAll(":", "").replaceAll("：", ""), right);
							}
						}
					}
				}
			} catch (Exception e) {
				throw e;
			}finally{
				finishCount_detail++;
			}
		}
	}
	
}
