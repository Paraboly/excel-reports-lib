package com.paraboly.reportlib.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * Need to add builder
 */
public class StyleUtils {
	private static Font getBoldFont(Sheet sheet) {
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 11);
		font.setFontName("Times New Roman");
		return font;
	}

	private static Font getHeaderFont(Sheet sheet) {
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Times New Roman");
		return font;
	}

	public static CellStyle getHeaderStyle(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(getBoldFont(sheet));
		return cellStyle;
	}

	public static CellStyle getBorderedBoldCellStyle(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFont(getBoldFont(sheet));
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0.00"));
		return cellStyle;
	}

	public static CellStyle getBorderedBoldCellStyleWithBackgroundColor(Sheet sheet, short bg) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(bg);
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setFont(getBoldFont(sheet));
		return cellStyle;
	}

	public static CellStyle getHeaderRowStyle(Sheet sheet) {
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.DOUBLE);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        cellStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
		cellStyle.setFont(getHeaderFont(sheet));
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0.00"));
		return cellStyle;
	}

	public static void setCurrency(Sheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0.00"));
	}

	public static void setCount(Sheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("#,##0"));
	}

	public static void setYear(Sheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat("0"));
	}

	public static void setPercentage(Sheet sheet, CellStyle cellStyle) {
		DataFormat format = sheet.getWorkbook().createDataFormat();
		cellStyle.setDataFormat(format.getFormat(BuiltinFormats.getBuiltinFormat(10)));
	}
}
