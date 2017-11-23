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

	static int pageSize = -1;//��ȡ����ҳ�����-1������ȡ����
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
		
		System.out.println("���������ӵ�ַ��");
		Scanner scan = new Scanner(System.in);
		htmlUrl = scan.nextLine();
		
		if(!"".equals(htmlUrl)){
			getNewHouse(htmlUrl,htmlUrl,houseList);
			
			//�����߳�ִ�����д��Excel�ļ�
			while(count != finishCount_detail || count != finishCount_dt || count != finishCount_hx || count != finishCount_kpxq){
				try {
					//System.out.println("�ȴ�������ҳ��ȡ��ɣ�");
					Thread.sleep(2000);
				} catch (InterruptedException e) {
				}
			}
			String fileName = exportExcel(houseList);
			System.out.println("===============�ɹ������ļ���" + fileName + "===============");
			System.out.println("===============ִ�����===============");
		}
	}

	/**
	 * ����Excel�ļ�
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
			//�����ļ���
			fileName = "newhouse_" + new SimpleDateFormat("yyyyMMddHHmmssS").format(new Date())+".xlsx";
			//String src = "blank.xlsx";
			
			destFile = new File(String.format("%s%s%s", "d:\\" , File.separator , fileName));
			//FileUtils.copyFile(new File(src), destFile);
			
			//����Excel����
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
			
			//------------------------------------------------д��ϸ��������------------------------------------------------//
			Sheet mxSheet = wb.createSheet();
			wb.setSheetName(0, "newhouse");

			List<String> mxHeads = new ArrayList<String>();//��ϸ�����ͷ����
			mxHeads.add("����");
			mxHeads.add("����");
			mxHeads.add("�۸�");
			
			mxHeads.add("��ҵ���");
			mxHeads.add("��Ȩ����");
			mxHeads.add("������");
			mxHeads.add("¥�̵�ַ");
			mxHeads.add("ռ�����");
			
			mxHeads.add("�ݻ���");
			mxHeads.add("�������");
			mxHeads.add("¥������");
			mxHeads.add("�ܻ���");
			mxHeads.add("ͣ��λ");
			
			mxHeads.add("����");
			mxHeads.add("��ҵ��");
			mxHeads.add("��ҵ��˾");
			mxHeads.add("¥��״��");
			
			//mxHeads.add("¥�̶�̬����");
			mxHeads.add("¥�̶�̬����");
			mxHeads.add("����ʱ��");
			mxHeads.add("��������");
			mxHeads.add("����ʱ��");
			
			mxHeads.add("��Ŀ���");
			
			mxHeads.add("��ַ");
			
			//д��ͷ
			Row mxRow = mxSheet.createRow(0);
			for(int i = 0;i< mxHeads.size();i++){
				mxSheet.setColumnWidth(i, 4500);
				ExcelUtil.writeCell(mxRow,i,Cell.CELL_TYPE_STRING,mxHeads.get(i),styleMap.get("headStyle"));
			}
			
			//д����������
			for(int i = 0;i<houseList.size();i++){
				Map<String,String> data = houseList.get(i);
				Row dataRow = mxSheet.createRow(mxSheet.getLastRowNum() + 1);
				int j = 0;
				for(int z = 0;z<mxHeads.size();z++){
					if("��ַ".equals(mxHeads.get(z))){
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),linkStyle);
						Cell cell = dataRow.getCell(j-1);
						cell.setCellType(Cell.CELL_TYPE_FORMULA);
						cell.setCellFormula("HYPERLINK(\"" + data.get(mxHeads.get(z)) + "\",\"" + data.get(mxHeads.get(z)) + "\")");
					}/*else if("����".equals(mxHeads.get(z))){
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),wrapStyle);
					}*/else{
						ExcelUtil.writeCell(dataRow,j++,Cell.CELL_TYPE_STRING,data.get(mxHeads.get(z)),styleMap.get("stringStyle"));
					}
				}
			}
			
			//дExcel�ļ�
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
	 * ��ȡ¥����Ϣ
	 * @param htmlUrl
	 * @param htmlUrl2
	 */
	public static void getNewHouse(String baseUrl, String htmlUrl,List<Map<String,String>> houseList) {
		System.out.println("=========����ִ�е�[" + page + "]ҳ=========");
		try {
			//��ȡ��ǰҳ��¥����Ϣ
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
				//һЩ�ϵ�������ҳ,�硰÷�ݡ��ȶ��߳���
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
			
			//�ж��Ƿ�����һҳ(ɸѡ���е�����ҳ����Щ��ҳû��)
			if(nhouse_list.size() > 0){
				//�°���ҳ
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
					//ҳ������ҳ
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
				//�ϰ���ҳ
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
	 * ��ת����¥�����顱ҳ��
	 */
	public static void directionToHouseDetail(String houseLink, String houseName, Map<String, String> houseMap){
		Document detailDoc_1  = getDoc(houseLink);
		if(null != detailDoc_1){
			//¥����ϸҳ��
			Elements navleft = detailDoc_1.getElementsByClass("navleft");
			String detailLink = "";//¥������
			String dtLink  = "";//¥�̶�̬
			String hxLink = "";//����
			String kpxqLink = "";//��������
			if(null != navleft && navleft.size() > 0){
				Elements navleft_a = navleft.get(0).getElementsByTag("a");
				for(int x = 0;x<navleft_a.size();x++){
					String linkName = navleft_a.get(x).text();
					if(linkName.indexOf("¥������") != -1 || linkName.indexOf("��ϸ��Ϣ") != -1){
						detailLink = navleft_a.get(x).attr("href");
						dtLink = detailLink.replace("housedetail.htm", "dongtai.htm");
						kpxqLink = detailLink.replace("housedetail.htm", "sale_history.htm");
					}
					if(linkName.indexOf("����") != -1){
						hxLink = navleft_a.get(x).attr("href");
					}
				}
			}
			houseMap.put("����", houseName);
			houseMap.put("��ַ", detailLink);
			
			if(!"".equals(detailLink)){
				count++;
				ExecuteThread e1 = new ExecuteThread(detailLink, houseName,"¥������",houseMap);
				e1.start();
				ExecuteThread e2 = new ExecuteThread(dtLink, houseName,"¥�̶�̬",houseMap);
				e2.start();
				ExecuteThread e3 = new ExecuteThread(hxLink, houseName,"����",houseMap);
				e3.start();
				ExecuteThread e4 = new ExecuteThread(kpxqLink, houseName,"��������",houseMap);
				e4.start();
				
				if(houseMap.size() > 0){
					houseList.add(houseMap);
				}
			}
		}
	}
	
	/**
	 * ����url����ȡjsoup��document
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
	 * ʹ��htmlclient��ȡhtml
	 * @return
	 */
	public static String getHtml(String htmlUrl) {
		// HttpClient ��ʱ����
		RequestConfig globalConfig = RequestConfig.custom()
				.setCookieSpec(CookieSpecs.STANDARD)
				.setConnectionRequestTimeout(timeOut)
				.setConnectTimeout(timeOut).build();
		CloseableHttpClient httpClient = HttpClients.custom()
				.setDefaultRequestConfig(globalConfig).build();

		// ����һ��GET����
		HttpGet httpGet = new HttpGet(htmlUrl);
		httpGet.addHeader(
				"User-Agent",
				"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.152 Safari/537.36");
		httpGet.addHeader(
				"Cookie",
				"_gat=1; nsfw-click-load=off; gif-click-load=on; _ga=GA1.2.1861846600.1423061484");
		CloseableHttpResponse response = null;
		try {
			// �������󣬲�ִ��
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
	 * ת��String
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
			System.out.println("=========������ȡ[" + houseName + "][" + type + "][" + houseLink + "]=========");
			try {
				if("¥������".equals(this.type)){
					getHouseDetail(houseLink, houseMap);
				}
				if("¥�̶�̬".equals(this.type)){
					getHouseDt(houseLink, houseMap);
				}
				if("����".equals(this.type)){
					getHx(houseLink, houseMap);
				}
				if("��������".equals(this.type)){
					getKpxq(houseLink, houseMap);
				}
			} catch (Exception e) {
				//e.printStackTrace();
				System.out.println("=========��ȡ[" + houseName + "][" + type + "][ʧ��]��" + e.getMessage() + "=========");
			}
		}

		/**
		 * ��ȡ��������
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
								houseMap.put("��������", kpjjluStr);
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
		 * ��ȡ����
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
						houseMap.put("����", hxStr);
					}
				}
			} catch (Exception e) {
				throw e;
			} finally {
				finishCount_hx++;
			}
		}
		
		/**
		 * ��ȡ��̬
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
							//�ɰ�¥�̶�̬
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
									houseMap.put("¥�̶�̬����", title);
									String dtHref = r_content.get(0).getElementsByTag("a").get(0).attr("href");
									Thread.sleep(1000);
									Document dtDoc_2 = getDoc(dtHref);
									if(null != dtDoc_2){
										Elements atc_wrapper = dtDoc_2.getElementsByClass("atc-wrapper");
										if(atc_wrapper.size() > 0){
											houseMap.put("¥�̶�̬����", atc_wrapper.get(0).text());
										}
									}
								}
								break;
							}
						}else{
							//ʹ���°�¥�̶�̬
							if("storyList".equals(dtli.get(i).attr("class"))){
								Elements storyList_a = dtli.get(i).getElementsByTag("a");
								if(storyList_a.size() > 0){
									String title = "";
									try {
										title = dtli.get(i).getElementsByClass("sLTime").get(0).text() + ":" + storyList_a.get(0).text();
									} catch (Exception e) {
									}
									if(!"".equals(title)){
										houseMap.put("¥�̶�̬����", title);
										String dtHref = storyList_a.get(0).attr("href");
										Thread.sleep(1000);
										Document dtDoc_2 = getDoc(dtHref);
										if(null != dtDoc_2){
											Elements atc_wrapper = dtDoc_2.getElementsByClass("atc-wrapper");
											if(atc_wrapper.size() > 0){
												houseMap.put("¥�̶�̬����", atc_wrapper.get(0).text());
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
		 * ��ȡ¥�������¥�̶�̬
		 * @param navleft_a
		 * @param houseMap
		 * @throws Exception 
		 */
		private void getHouseDetail(String detailLink,
				Map<String, String> houseMap) throws Exception {
			try {
				houseMap.put("����", houseName);
				houseMap.put("��ַ", detailLink);
				
				//��ȡ����ϸ��Ϣ���е�����
				//Document detailDoc_2  = Jsoup.connect(detailLink).timeout(timeOut).get();
				Thread.sleep(1000);
				Document detailDoc_2 = getDoc(detailLink);
				if(null != detailDoc_2){
					Elements h1_label = detailDoc_2.getElementsByClass("h1_label");
					for(int i = 0 ;i<h1_label.size();i++){
						String biemin = h1_label.get(i).text();
						if(null != biemin && biemin.indexOf("����") != -1){
							houseMap.put("����", biemin);
							break;
						}
					}
					
					Elements main_left = detailDoc_2.getElementsByClass("main-left");
					if(null != main_left && main_left.size() > 0){
						//�۸�
						Elements main_info_price = main_left.get(0).getElementsByClass("main-info-price");
						if(null != main_info_price && main_info_price.size() > 0){
							houseMap.put("�۸�", main_info_price.get(0).text());
						}
						//��Ŀ���
						Elements intro = main_left.get(0).getElementsByClass("intro");
						if(intro.size() > 0){
							houseMap.put("��Ŀ���", intro.get(0).text());
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
								houseMap.put(left.replaceAll(":", "").replaceAll("��", ""), right);
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
