package com.paraboly.reportlib.utils;

import com.paraboly.reportlib.GenericReports;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Need to add builder
 */
public class StyleUtils {
	private static Font getBoldFont(XSSFSheet sheet, int size) {
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) size);
		font.setFontName("Times New Roman");
		return font;
	}

	private static Font getTitleBoldFont(XSSFSheet sheet, int size) {
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) size);
		font.setFontName("Times New Roman");
		return font;
	}

	private static Font getHeaderFont(XSSFSheet sheet, int size) {
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) size);
		font.setFontName("Times New Roman");
		return font;
	}
	public static CellStyle getTitleHeaderStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setFont(getTitleBoldFont(sheet,size));
		return cellStyle;
	}



	public static CellStyle getHeaderStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setFont(getBoldFont(sheet, size));
		return cellStyle;
	}

	public static CellStyle getBorderedCellStyle(XSSFSheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		return cellStyle;
	}

	public static CellStyle getBorderedBoldCellStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setFont(getBoldFont(sheet, size));
		return cellStyle;
	}

	public static CellStyle getBottomTitleCellStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setFont(getBoldFont(sheet, size));
		return cellStyle;
	}

	public static CellStyle getBorderedBoldCurrencyCellStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setFont(getBoldFont(sheet,size));
		return cellStyle;
	}

	public static CellStyle getBorderedBoldCellStyleWithBackgroundColor(XSSFSheet sheet, short bg, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(bg);
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setFont(getBoldFont(sheet,size));
		return cellStyle;
	}

	public static CellStyle getHeaderRowStyle(XSSFSheet sheet, int size) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setFont(getHeaderFont(sheet,size));
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		return cellStyle;
	}

	public static void setCurrency(XSSFSheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0.00\\ TL"));
	}

	public static void setCount(XSSFSheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0"));
	}

	public static void setYear(XSSFSheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("0"));
	}

	public static void setText(XSSFSheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat(""));
	}
	public static void setLink(XSSFSheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat(""));
	}

	public static void setPercentage(XSSFSheet sheet, CellStyle cellStyle, GenericReports.ColumnMetadata columnMetadata) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		if(columnMetadata.getDecimalPoint()==0){
			cellStyle.setDataFormat(format.getFormat("% 0"));
		}else{
			cellStyle.setDataFormat(format.getFormat("% 0.0"));
		}

	}
}
