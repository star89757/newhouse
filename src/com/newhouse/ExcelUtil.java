package com.newhouse;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;

/**
 *
 * ex_gz_zhangxing03
 */
public class ExcelUtil {

	/**
	 * 写入单元格值(Excel2007)
	 * @param row 
	 * @param colIndex
	 * @param cellTypeString
	 * @param value
	 */
	public static void writeCell(XSSFRow row, int colIndex, int cellTypeString, Object value) {
		XSSFCell cell = row.createCell(colIndex);
		cell.setCellType(cellTypeString);//文本格式
		if(cellTypeString == XSSFCell.CELL_TYPE_FORMULA){
			cell.setCellFormula((String) value);
		}else{
			if(value instanceof Boolean){
				cell.setCellValue((Boolean)value);
			}else if(value instanceof Calendar){
				cell.setCellValue((Calendar)value);
			}else if(value instanceof Date){
				cell.setCellValue((Date)value);
			}else if(value instanceof Double){
				cell.setCellValue((Double)value);
			}else if(value instanceof BigDecimal){
				cell.setCellValue(((BigDecimal)value).doubleValue());
			}else if(value instanceof RichTextString){
				cell.setCellValue((RichTextString)value);
			}else if(value instanceof Integer){
				cell.setCellValue((Integer)value);
			}else{
				cell.setCellValue((String)value);
			}
		}
	}
	
	/**
	 * 写入单元格值(Excel2007)
	 * @param row
	 * @param colIndex
	 * @param cellTypeString
	 * @param value
	 * @param style
	 */
	public static void writeCell(XSSFRow row, int colIndex, int cellTypeString, Object value,XSSFCellStyle style) {
		XSSFCell cell = row.createCell(colIndex);
		cell.setCellType(cellTypeString);//文本格式
		if(cellTypeString == XSSFCell.CELL_TYPE_FORMULA){
			cell.setCellFormula((String) value);
		}else{
			if(value instanceof Boolean){
				cell.setCellValue((Boolean)value);
			}else if(value instanceof Calendar){
				cell.setCellValue((Calendar)value);
			}else if(value instanceof Date){
				cell.setCellValue((Date)value);
			}else if(value instanceof Double){
				cell.setCellValue((Double)value);
			}else if(value instanceof BigDecimal){
				cell.setCellValue(((BigDecimal)value).doubleValue());
			}else if(value instanceof RichTextString){
				cell.setCellValue((RichTextString)value);
			}else if(value instanceof Integer){
				cell.setCellValue((Integer)value);
			}else{
				cell.setCellValue((String)value);
			}
		}
		if(style != null){
			cell.setCellStyle(style);
		}
	}
	
	/**
	 * 写SXSSFWorkbook大数据
	 * @param row
	 * @param colIndex
	 * @param cellTypeString
	 * @param value
	 */
	public static void writeCell(Row row, int colIndex, int cellTypeString, Object value) {
		Cell cell = row.createCell(colIndex);
		cell.setCellType(cellTypeString);//文本格式
		if(cellTypeString == XSSFCell.CELL_TYPE_FORMULA){
			cell.setCellFormula((String) value);
		}else{
			if(value instanceof Boolean){
				cell.setCellValue((Boolean)value);
			}else if(value instanceof Calendar){
				cell.setCellValue((Calendar)value);
			}else if(value instanceof Date){
				cell.setCellValue((Date)value);
			}else if(value instanceof Double){
				cell.setCellValue((Double)value);
			}else if(value instanceof BigDecimal){
				cell.setCellValue(((BigDecimal)value).doubleValue());
			}else if(value instanceof RichTextString){
				cell.setCellValue((RichTextString)value);
			}else if(value instanceof Integer){
				cell.setCellValue((Integer)value);
			}else{
				cell.setCellValue((String)value);
			}
		}
	}
	
	/**
	 * 写SXSSFWorkbook大数据
	 * @param row
	 * @param colIndex
	 * @param cellTypeString
	 * @param value
	 * @param style
	 */
	public static void writeCell(Row row, int colIndex, int cellTypeString,
			Object value, CellStyle style) {
		Cell cell = row.createCell(colIndex);
		cell.setCellType(cellTypeString);//文本格式
		if(cellTypeString == XSSFCell.CELL_TYPE_FORMULA){
			cell.setCellFormula((String) value);
		}else{
			if(value instanceof Boolean){
				cell.setCellValue((Boolean)value);
			}else if(value instanceof Calendar){
				cell.setCellValue((Calendar)value);
			}else if(value instanceof Date){
				cell.setCellValue((Date)value);
			}else if(value instanceof Double){
				cell.setCellValue((Double)value);
			}else if(value instanceof BigDecimal){
				cell.setCellValue(((BigDecimal)value).doubleValue());
			}else if(value instanceof RichTextString){
				cell.setCellValue((RichTextString)value);
			}else if(value instanceof Integer){
				cell.setCellValue((Integer)value);
			}else{
				cell.setCellValue((String)value);
			}
		}
		if(style != null){
			cell.setCellStyle(style);
		}
	}
	
	/**
	 * 写公式单元格值
	 * @param row
	 * @param colIndex
	 * @param value
	 * @param style
	 */
	public static void writeCellFORMULA(XSSFRow row, int colIndex, Object value,XSSFCellStyle style) {
		XSSFCell cell = row.createCell(colIndex);
		cell.setCellFormula((String) value);
		if(style != null){
			cell.setCellStyle(style);
		}
	}
	

	/**
	 * 写入单元格值(Excel2003)
	 * @param row
	 * @param colIndex
	 * @param cellTypeString
	 * @param value
	 */
	public static void writeCell(HSSFRow row, int colIndex, int cellTypeString,
			String value) {
		HSSFCell cell = row.createCell(colIndex);            		
		cell.setCellType(cellTypeString);//文本格式
		cell.setCellValue(value);//写入内容
	}  
	
	/**
	 * 
	 * 获取Cell的值
	 * 
	 * @param cell cell对象
	 * @return
	 */
	public static Object getCellValue(Cell cell) {
        String ret = "";  
        if(null != cell){
	        switch (cell.getCellType()) {  
	        case Cell.CELL_TYPE_BLANK:  
	            ret = "";  
	            break;  
	        case Cell.CELL_TYPE_BOOLEAN:  
	            ret = String.valueOf(cell.getBooleanCellValue());  
	            break;  
	        case Cell.CELL_TYPE_ERROR:  
	            ret = null;  
	            break;  
	        case Cell.CELL_TYPE_FORMULA:  
	        	if(cell.getCachedFormulaResultType() == 0){
	        		if (DateUtil.isCellDateFormatted(cell)) {
	        			Date date = cell.getDateCellValue();
	        			return date;
	        		}else{
	        			ret = String.valueOf(cell.getNumericCellValue());
	        		}
	        	};
	        	if(cell.getCachedFormulaResultType() == 1){
	        		ret = cell.getStringCellValue();
	        	}
	            break;  
	        case Cell.CELL_TYPE_NUMERIC:  
	            if (DateUtil.isCellDateFormatted(cell)) {   
	                Date theDate = cell.getDateCellValue();  
	                return theDate;  
	            } else {   
	                ret = NumberToTextConverter.toText(cell.getNumericCellValue());  
	            }  
	            break;  
	        case Cell.CELL_TYPE_STRING:  
	            ret = cell.getRichStringCellValue().getString();  
	            break;  
	        default:  
	            ret = "";  
	        }  
        }else{
        	ret = "";
        }
        return ret; 
    }
	
	/**
	 * 获取double值
	 * @param value
	 * @return
	 */
	public static Double getDoubleValue(Object value){
		if("".equals(value) || null == value){
			return 0.0;
		}
		if(value instanceof String){
			return Double.parseDouble((String)value);
		}else if(value instanceof BigDecimal){
			return ((BigDecimal)value).doubleValue();
		}else{
			return Double.parseDouble(value.toString());
		}
	}
	
	/**
	 * 获取int值
	 * @param value
	 * @return
	 */
	public static Integer getIntValue(Object value){
		if("".equals(value) || null == value){
			return 0;
		}
		if(value instanceof String){
			return Integer.parseInt((String)value);
		}else if(value instanceof BigDecimal){
			return ((BigDecimal)value).intValue();
		}else{
			return Integer.parseInt(value.toString());
		}
	}
	
	/**
	 * 将Excel列符转换为数字
	 * @param colStr
	 * @param length
	 * @return
	 */
	public static int excelColStrToNum(String colStr,int length){
		int num = 0;
		int result = 0;
		for(int i = 0;i<length;i++){
			char ch = colStr.charAt(length - i - 1);
			num = (int)(ch - 'A' + 1);
			num*=Math.pow(26, i);
			result += num;
		}
		return result;
	}
	
	/**
	 * 获取一些基本样式
	 * 1.加粗的报表标题字体样式
	 * 2.普通的字符串样式
	 * 3.带普通边框的，表头加粗字体样式
	 * 4.带普通边框的，数值样式（正数正常字体，负数红色字体）
	 * 5.带普通边框，灰色背景色的数值汇总样式
	 * @param wb
	 * @return
	 */
	public static Map<String,CellStyle> getCommonStyle(Workbook wb){
		Map<String,CellStyle> map = new HashMap<String, CellStyle>();

		CellStyle style  = wb.createCellStyle();
		style.setBottomBorderColor(HSSFColor.BLACK.index);
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		
		CellStyle stringStyle = wb.createCellStyle();//(CellStyle) style.clone();
		stringStyle.cloneStyleFrom(style);
		//stringStyle.setWrapText(true);//自动换行
		stringStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		
		//报表名称
		Font titlef = wb.createFont();
		titlef.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		titlef.setFontHeightInPoints((short) 18);
		CellStyle titleStyle = wb.createCellStyle();
		titleStyle.setFont(titlef);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		//titleStyle.setWrapText(true);
		
		Font redf = wb.createFont();
		redf.setFontHeightInPoints((short) 10);
		redf.setColor(HSSFColor.RED.index);
		
		Font normalf = wb.createFont();
		normalf.setFontHeightInPoints((short) 10);
		
		Font headf = wb.createFont();
		headf.setFontHeightInPoints((short) 11);
		headf.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		
		CellStyle negativeNumStyle = wb.createCellStyle();
		negativeNumStyle.cloneStyleFrom(style);
		negativeNumStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		negativeNumStyle.setFont(redf);
		negativeNumStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
		
		CellStyle notNegativeNumStyle = wb.createCellStyle();
		notNegativeNumStyle.cloneStyleFrom(style);
		notNegativeNumStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		notNegativeNumStyle.setFont(normalf);
		notNegativeNumStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
		
		CellStyle headStyle = (CellStyle) wb.createCellStyle();
		headStyle.cloneStyleFrom(style);
		headStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headStyle.setFont(headf);
		headStyle.setWrapText(true);
		
		CellStyle sumStyleStr = (CellStyle) wb.createCellStyle();
		sumStyleStr.cloneStyleFrom(style);
		sumStyleStr.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		sumStyleStr.setFillPattern(CellStyle.SOLID_FOREGROUND);
		sumStyleStr.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		
		CellStyle sumStyleNum =  wb.createCellStyle();
		sumStyleNum.cloneStyleFrom(notNegativeNumStyle);
		sumStyleNum.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		sumStyleNum.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		map.put("stringStyle", stringStyle);
		map.put("negativeNumStyle", negativeNumStyle);
		map.put("notNegativeNumStyle", notNegativeNumStyle);
		map.put("headStyle", headStyle);
		map.put("titleStyle", titleStyle);
		map.put("sumStyleStr", sumStyleStr);
		map.put("sumStyleNum", sumStyleNum);
		
		return map;
	}
	
	/**
	 * 获取常用字体
	 * headFontNormal:标题字体
	 * headFontRed：红色标题字体
	 * normalf：正常字体
	 * redf：红色正常字体
	 * @param wb
	 * @return
	 */
	public static Map<String,Font> getCommonFont(Workbook wb){
		Map<String,Font> map = new HashMap<String, Font>();

		Font redf = wb.createFont();
		redf.setFontHeightInPoints((short) 10);
		redf.setColor(HSSFColor.RED.index);
		
		Font normalf = wb.createFont();
		normalf.setFontHeightInPoints((short) 10);
		
		Font headf = wb.createFont();
		headf.setFontHeightInPoints((short) 11);
		headf.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		
		Font headfr = wb.createFont();
		headfr.setFontHeightInPoints((short) 11);
		headfr.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		headfr.setColor(HSSFColor.RED.index);
		
		map.put("headFontNormal", headf);
		map.put("headFontRed", headfr);
		map.put("normalFont", normalf);
		map.put("redFont", redf);
		
		return map;
	}
	
	/**
	 * 将数字转换Excel列符
	 * @param columnIndex
	 * @return
	 */
	public static String excelColIndexToStr(int columnIndex){
		if(columnIndex < 0){
			return null;
		}
		String columnStr = "";
		columnIndex--;
		do {
			if(columnStr.length()>0){
				columnIndex--;
			}
			columnStr = ((char)(columnIndex %26 + (int)'A')) + columnStr;
			columnIndex = (int)((columnIndex - columnIndex%26)/26);
		} while (columnIndex > 0);
		return columnStr;
	}
	
	public static void main(String[] args) {
		System.out.println(excelColIndexToStr(1));
	}

}
