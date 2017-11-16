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

import org.apache.commons.io.FileUtils;
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
	static int finishCount = 0;
	static List<Map<String,String>> houseList = new ArrayList<Map<String,String>>();
	
	
	public static void main(String[] args) {
		
		String htmlUrl = "http://newhouse.dg.fang.com/house/s";
		getNewHouse(htmlUrl,htmlUrl,houseList);
		
		//�����߳�ִ�����д��Excel�ļ�
		while(count != finishCount){
			try {
				//System.out.println("�ȴ�������ҳ��ȡ��ɣ�");
				Thread.sleep(2000);
			} catch (InterruptedException e) {
			}
		}
		/*for(int i =0;i<houseList.size();i++){
			System.out.println(houseList.get(i).toString());
		}*/
		exportExcel(houseList);
		System.out.println("===============ִ�����===============");
	}

	/**
	 * ����Excel�ļ�
	 * @param excelFile
	 * @param houseList
	 */
	@SuppressWarnings("deprecation")
	private static void exportExcel(List<Map<String, String>> houseList) {
		
		FileOutputStream out = null;
		String fileName = "";
		
		try {
			//�����ļ���
			fileName = "newhouse_" + new SimpleDateFormat("yyyyMMddHHmmssS").format(new Date())+".xlsx";
			String src = new File(GetNewHouse.class.getResource("/").getPath()).getParent() + File.separator + "tmpl" + File.separator + "blank.xlsx";
			
			File destFile = new File(String.format("%s%s%s", "d:\\" , File.separator , fileName));
			FileUtils.copyFile(new File(src), destFile);
			
			//����Excel����
			Workbook wb = new SXSSFWorkbook(10000);
			Map<String, CellStyle> styleMap = ExcelUtil.getCommonStyle(wb);
			CellStyle linkStyle  = wb.createCellStyle();
			linkStyle.cloneStyleFrom(styleMap.get("stringStyle"));
			
			Font f = wb.createFont();
			f.setUnderline(Font.U_SINGLE);
			f.setColor(HSSFColor.BLUE.index);
			linkStyle.setFont(f);
			
			//------------------------------------------------д��ϸ��������------------------------------------------------//
			Sheet mxSheet = wb.createSheet();
			wb.setSheetName(0, "newhouse");

			List<String> mxHeads = new ArrayList<String>();//��ϸ�����ͷ����
			mxHeads.add("����");
			
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
			
			mxHeads.add("����μ�����");
			mxHeads.add("��ҵ��");
			mxHeads.add("��ҵ��˾");
			mxHeads.add("¥��״��");
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
					}else{
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
	}

	/**
	 * ��ȡ������Ϣ
	 * @param htmlUrl
	 * @param htmlUrl2
	 */
	public static void getNewHouse(String baseUrl, String htmlUrl,List<Map<String,String>> houseList) {
		System.out.println("=========����ִ�е�[" + page + "]ҳ=========");
		try {
			//��ȡ��ǰҳ�ķ�����Ϣ
			//Document doc  = Jsoup.connect(htmlUrl).timeout(timeOut).get();
			Document doc = getDoc(htmlUrl);
			if(null == doc){
				return;
			}
			Elements nhouse_list = doc.getElementsByClass("nhouse_list");
			for(int i = 0;i<nhouse_list.size();i++){
				Elements nhouse_li = nhouse_list.get(i).getElementsByTag("li");
				for(int j = 0;j<nhouse_li.size();j++){
					Elements nhouse_li_a = nhouse_li.get(j).getElementsByTag("a");
					if(nhouse_li_a.size() > 1){
						String houseName = nhouse_li_a.get(1).text();
						String houseLink = nhouse_li_a.get(1).attr("href");
						//Document detailDoc_1  = Jsoup.connect(houseLink).timeout(timeOut).get();
						Thread.sleep(1000);
						//�½����߳�ȥ�����ݡ���ϸ��Ϣ��
						count++;
						ExecuteThread e = new ExecuteThread(houseLink, houseName);
						e.start();
					}
				}
			}
			
			//�ж��Ƿ�����һҳ
			Elements otherpage = doc.getElementsByClass("otherpage");
			if(otherpage.size()>0){
				Elements otherpage_a = otherpage.get(0).getElementsByTag("a");
				for(int i = 0;i<otherpage_a.size();i++){
					if(">".equals(otherpage_a.get(i).text())){
						if(-1 == pageSize){
							page++;
							getNewHouse(baseUrl, baseUrl + otherpage_a.get(i).attr("href"),houseList);
						}else{
							if(page < pageSize){
								page++;
								getNewHouse(baseUrl, baseUrl + otherpage_a.get(i).attr("href"),houseList);
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
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
		try {
			// �������󣬲�ִ��
			CloseableHttpResponse response = httpClient.execute(httpGet);
			InputStream in = response.getEntity().getContent();
			String html = convertStreamToString(in);
			Document doc = Jsoup.parse(html);
			return doc.html();
		} catch (Exception e) {
			e.printStackTrace();
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
		
		public ExecuteThread(String url,String name){
			this.houseLink = url;
			this.houseName = name;
		}
		
		public void run(){
			System.out.println("=========������ȡ[" + houseLink + "]=========");
			try {
				//��ʼ����
				Document detailDoc_1  = getDoc(houseLink);
				if(null != detailDoc_1){
					//������ϸҳ��
					Elements navleft = detailDoc_1.getElementsByClass("navleft");
					if(null != navleft && navleft.size() > 0){
						Elements navleft_a = navleft.get(0).getElementsByTag("a");
						for(int x = 0;x<navleft_a.size();x++){
							String linkName = navleft_a.get(x).text();
							if(linkName.indexOf("¥������") != -1 || linkName.indexOf("��ϸ��Ϣ") != -1){
								String detailLink = navleft_a.get(x).attr("href");
								Map<String,String> houseMap = new HashMap<String,String>();
								houseMap.put("����", houseName);
								houseMap.put("��ַ", detailLink);
								
								//��ȡ����ϸ��Ϣ���е�����
								//Document detailDoc_2  = Jsoup.connect(detailLink).timeout(timeOut).get();
								Thread.sleep(1000);
								Document detailDoc_2 = getDoc(detailLink);
								if(null == detailDoc_2){
									continue;
								}
								Elements main_left = detailDoc_2.getElementsByClass("main-left");
								if(null != main_left && main_left.size() > 0){
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
								houseList.add(houseMap);
								break;
							}
						}
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}finally{
				finishCount++;
			}
		}
	}
	
}
