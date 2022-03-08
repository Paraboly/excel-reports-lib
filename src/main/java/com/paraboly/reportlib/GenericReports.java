package com.paraboly.reportlib;

import com.sun.corba.se.spi.orbutil.threadpool.Work;
import lombok.Data;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xddf.usermodel.XDDFEffectContainer;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.jfree.chart.labels.PieSectionLabelGenerator;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot3D;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment;

import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.paraboly.reportlib.utils.StyleUtils.*;

// build

public class GenericReports {
	@Data
	public static class ReportData {
		private List<?> elementList;
		private LinkedHashMap<String, ColumnMetadata> columnToMetadataMapping;
		private String reportType;
		private int fontSize=12;
		private int headerFontSize=12;
		private int titleFontSize=14;
		private Integer year;
		private Integer headerStartOffsetX;
		private Integer headerEndOffsetX;
		private Integer headerStartOffsetY;
		private Integer headerEndOffsetY;
		private String biddingDepartment;
		private String biddingType;
		private String biddingProcedure;
		private LinkedList<ChartProps> chartPropsLinkedList;
		private ChartProps chartProps;
		private LinkedList<String> addToTotalSumList;
		private String totalSumTitle;
		private Boolean disableBottomRow = false;
		private String rowColorFunction;
		private Integer yearCount;
		private Boolean mergeTwoRow = false;
	}

	@Data
	public static class ColumnMetadata {
		private String functionName;
		private Function customFunction;
		private Integer columnSize = 1;
		private String bottomCalculation = "string:"; // potential values are sum, avg, or string:BOTTOM_NAME
		private String bottomCalculationText = "";
		private String bottomValue;
		private String bottomTitle;
		private String cellContent = "text"; // potential values are money, percentage, count, year
		private String alignment = "CENTER";
		private Boolean isDiscount=false;
		private Integer decimalPoint=999;
		private Boolean isMerged = false;
	}

	@Data
	public static class ChartProps {
		private String groupFunctionName;
		private String groupLabel;
		private String[] valueFunctionName;
		private String[] valueLabel;
		private String type;
		private String title;
		private String groupKey;
		private String[] valueKey;
	}

	public static class Builder {
		private static List<ReportData> reportDataList;
		private static String filename;
		private static XSSFWorkbook wb;

		public Builder(String filename) {
			this.filename = filename;
			reportDataList = new ArrayList<>();
			wb = new XSSFWorkbook();
		}

		public Builder addData(ReportData data) {
			reportDataList.add(data);
			return this;
		}

		public static XSSFWorkbook create() {
			for (ReportData reportData: reportDataList) {
				XSSFSheet sheet = wb.createSheet(reportData.getReportType());

				if(reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
					sheet.setZoom(60);
				}
				TableMapperExtended tableMapperExtended = getReportTable(reportData, sheet);
				tableMapperExtended.setStartOffsetX(0);
				tableMapperExtended.setStartOffsetY(0);
				tableMapperExtended.write(sheet, reportData);
				if (reportData.chartPropsLinkedList != null) {
					AtomicInteger i = new AtomicInteger(0);
					reportData.chartPropsLinkedList.forEach(chartProps -> {
						fillChartProps(chartProps, reportData.getColumnToMetadataMapping());
						tableMapperExtended.addChart(sheet, reportData.getElementList(), chartProps, i.getAndIncrement());
					});
				}
				sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
				sheet.setFitToPage(true);
				sheet.getPrintSetup().setFitWidth((short)1);
				sheet.getPrintSetup().setFitHeight((short)1);
			}
			return wb;
		}

		private static ChartProps fillChartProps(ChartProps chartProps, LinkedHashMap<String, ColumnMetadata> columnMetadata) {
			chartProps.setGroupFunctionName(columnMetadata.get(chartProps.getGroupKey()).getFunctionName());
			chartProps.setGroupLabel(chartProps.getGroupKey());

			chartProps.setValueFunctionName(new String[chartProps.valueKey.length]);

			for(int t = 0; t < chartProps.valueKey.length; t++){
				chartProps.getValueFunctionName()[t] = columnMetadata.get(chartProps.getValueKey()[t]).getFunctionName();
			}

			chartProps.setValueLabel(chartProps.getValueKey());
			return chartProps;
		}
		private static CellStyle getCellStyle(XSSFSheet sheet, String type, ColumnMetadata columnMetadata, int size){
			CellStyle dataStyle = getBorderedBoldCellStyle(sheet, size);
			CellStyle headerStyle = getHeaderRowStyle(sheet, size);
			CellStyle currStyle = getBorderedBoldCurrencyCellStyle(sheet,size);
			if(type.equals("year")){
				CellStyle yearStyle = sheet.getWorkbook().createCellStyle();
				yearStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					yearStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					yearStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setYear(sheet, yearStyle);
				return yearStyle;
			}else if(type.equals("money")){
				CellStyle currencyStyle = sheet.getWorkbook().createCellStyle();
				currencyStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					currencyStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					currencyStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setCurrency(sheet, currencyStyle);
				return currencyStyle;
			}else if(type.equals("percentage")){
				CellStyle percentageStyle = sheet.getWorkbook().createCellStyle();
				percentageStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					percentageStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					percentageStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setPercentage(sheet, percentageStyle, columnMetadata );
				return percentageStyle;
			}else if(type.equals("count")){
				CellStyle countStyle = sheet.getWorkbook().createCellStyle();
				countStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					countStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					countStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setCount(sheet, countStyle);
				return countStyle;
			}else if(type.equals("text")){
				CellStyle textStyle = sheet.getWorkbook().createCellStyle();
				textStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					textStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					textStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setText(sheet, textStyle);
				return textStyle;
			}
			else if(type.equals("link")){
				CellStyle linkStyle = sheet.getWorkbook().createCellStyle();
				linkStyle.cloneStyleFrom(dataStyle);
				if (columnMetadata.getAlignment().equals("RIGHT")){
					linkStyle.setAlignment(HorizontalAlignment.RIGHT);
				}else if(columnMetadata.getAlignment().equals("LEFT")){
					linkStyle.setAlignment(HorizontalAlignment.LEFT);
				}
				setLink(sheet, linkStyle);
				return linkStyle;
			}
			else{
				return headerStyle;
			}
		}


		private static TableMapperExtended getReportTable(ReportData reportData, XSSFSheet sheet) {
			LinkedHashMap<String, ColumnDefinition> map = new LinkedHashMap<>();
			CellStyle headerStyle = getHeaderRowStyle(sheet, reportData.headerFontSize);
			reportData.getColumnToMetadataMapping().forEach((columnName, columnMetadata) -> {
				CellStyle fieldStyle = null;
				switch (columnMetadata.getCellContent()) {
					case "money":
						fieldStyle = getCellStyle(sheet, "money", columnMetadata,reportData.fontSize);
						break;
					case "percentage":
						fieldStyle = getCellStyle(sheet, "percentage", columnMetadata,reportData.fontSize);
						break;
					case "count":
						fieldStyle = getCellStyle(sheet, "count", columnMetadata,reportData.fontSize);
						break;
					case "year":
						fieldStyle = getCellStyle(sheet, "year", columnMetadata,reportData.fontSize);
						break;
					case "text":
						fieldStyle = getCellStyle(sheet, "text", columnMetadata,reportData.fontSize);
						break;
					case "link":
						fieldStyle = getCellStyle(sheet, "link", columnMetadata,reportData.fontSize);
						break;
				}
				map.put(columnName,
						new ColumnDefinition<String>(
								columnMetadata.getColumnSize(), columnName, fieldStyle, headerStyle,
								columnMetadata.getBottomCalculation(),columnMetadata.getBottomCalculationText(), columnMetadata.getBottomValue(), columnMetadata.getBottomTitle(), reportData.getDisableBottomRow(), reportData, columnMetadata.getAlignment(), columnMetadata.getIsDiscount(), columnMetadata.getDecimalPoint(), columnMetadata.getIsMerged(), columnMetadata.cellContent));
			});

			for (Object data: reportData.getElementList()) {
				reportData.getColumnToMetadataMapping().forEach((columnName, columnMetadata) -> {
					try {
						map.get(columnName).getData().add(
								columnMetadata.getCustomFunction() == null ?
										data.getClass().getMethod(columnMetadata.getFunctionName()).invoke(data) :
										invokeCustomMethod(data, columnMetadata.getCustomFunction())
						);
					} catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
						e.printStackTrace();
					}
				});
			}

			return new TableMapperExtended(reportData.getReportType(), new ArrayList<>(map.values()), reportData);
		}

		private static Object invokeCustomMethod(Object data, Function function) {
			return function.apply(data);
		}
	}

	private static class ColumnDefinition<T> {
		private int columnSize;
		private String column;
		private List<T> data;
		private CellStyle columnStyle;
		private CellStyle headerStyle;
		private int offsetYCounter;
		private int startOffsetX;
		private int startOffsetY;
		private String bottomCalculation;
		private String bottomCalculationText;
		private String bottomValue;
		private String bottomTitle;
		private Boolean disableBottomRow;
		private ReportData reportData;
		private String alignment;
		private Boolean isDiscount;
		private Integer decimalPoint;
		private Boolean isMerged;
		private String cellContent;

		public ColumnDefinition(int columnSize,
								String column,
								CellStyle columnStyle,
								CellStyle headerStyle,
								String bottomCalculation,
								String bottomCalculationText,
								String bottomValue,
								String bottomTitle,
								Boolean disableBottomRow,
								ReportData reportData,
								String alignment,
								Boolean isDiscount,
								Integer decimalPoint,
								Boolean isMerged,
								String cellContent
		) {
			this.columnSize = columnSize;
			this.column = column;
			this.columnStyle = columnStyle;
			this.headerStyle = headerStyle;
			this.bottomCalculation = bottomCalculation;
			this.bottomCalculationText = bottomCalculationText;
			this.bottomValue = bottomValue;
			this.bottomTitle = bottomTitle;
			this.disableBottomRow = disableBottomRow;
			this.reportData = reportData;
			this.alignment = alignment;
			this.isDiscount = isDiscount;
			this.decimalPoint = decimalPoint;
			this.isMerged = isMerged;
			this.cellContent = cellContent;
			data = new ArrayList<T>();
		}

		public int getColumnSize() {
			return columnSize;
		}

		public List<T> getData() {
			return data;
		}

		public void setData(List<T> data) {
			this.data = data;
		}

		public int getOffsetYCounter() {
			return offsetYCounter;
		}

		public static void reformatCell(XSSFCell dataCell, int columnSize){
			String fontName = dataCell.getCellStyle().getFont().getFontName();
			int fontSize = dataCell.getCellStyle().getFont().getFontHeightInPoints();
			String value = dataCell.getStringCellValue();

			java.awt.Font font = new java.awt.Font(fontName, 0, fontSize);
			FontRenderContext frc = new FontRenderContext(new AffineTransform(), true, true);

			float textwidth = (float) font.getStringBounds(value, frc).getWidth() * 1.088f;
			float textheight = (float) font.getStringBounds(value, frc).getHeight() * 1.088f;


			float columnWidthInPoints = dataCell.getSheet().getDefaultColumnWidth() * 7.3f * columnSize;

			if(textwidth > columnWidthInPoints){

				float oneCharLength = textheight;
				float count = textwidth / columnWidthInPoints;

				float multiplier = (float) Math.ceil(count);
				float length = oneCharLength * multiplier;

				if(length > dataCell.getRow().getHeightInPoints())
					dataCell.getRow().setHeightInPoints(length * 1.5f);
			}
			else if(textheight * 1.5f > dataCell.getRow().getHeightInPoints()) {
				dataCell.getRow().setHeightInPoints(textheight * 1.5f);
			}
		}
		public void write(XSSFSheet sheet, int startOffsetY, int startOffsetX) {
			sheet.setDefaultColumnWidth(14);

			this.startOffsetX = startOffsetX;
			this.startOffsetY = startOffsetY;

			offsetYCounter = startOffsetY;
			Row columnHeaderRow = sheet.getRow(offsetYCounter);
			if(columnHeaderRow == null) {
				columnHeaderRow = sheet.createRow(offsetYCounter);
			}
			double height = 0;
			if (this.reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
				height = 17.0;
			}else if (this.reportData.reportType.substring(0,1).equals(" ")){
				height = 6.0;
			}
			else{
				height = 7.0;
			}


			columnHeaderRow.setHeight((short)height);
			columnHeaderRow.setHeightInPoints((4* columnHeaderRow.getHeight()));



			if(columnSize >= 1 && reportData.mergeTwoRow==true) {
				CellRangeAddress region = new CellRangeAddress(offsetYCounter, offsetYCounter+1, startOffsetX, startOffsetX + columnSize - 1);
				sheet.addMergedRegion(region);
				RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
				offsetYCounter += 2;

				sheet.getRow(columnHeaderRow.getRowNum() + 1).setHeight((short) height);
				sheet.getRow(columnHeaderRow.getRowNum() + 1)
						.setHeightInPoints(4 * sheet.getRow(columnHeaderRow.getRowNum() + 1).getHeight());


			}else if(columnSize==1 && reportData.mergeTwoRow==false){
				offsetYCounter += 1;
			}
			else if(columnSize>1 && reportData.mergeTwoRow==false){
				if(!isMerged){
					CellRangeAddress region = new CellRangeAddress(offsetYCounter, offsetYCounter, startOffsetX, startOffsetX + columnSize - 1);
					sheet.addMergedRegion(region);
					RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
				}
				offsetYCounter += 1;
			}
			if(!(isMerged && column.equals("İHALE TÜRÜ"))) { // control
				Cell cell = columnHeaderRow.createCell(startOffsetX);
				cell.setCellValue(column);
				if (headerStyle != null)
					cell.setCellStyle(headerStyle);
			}

			for (int i = 0; i <= data.size(); i++) {
				XSSFRow dataRow = sheet.getRow(i + offsetYCounter);
				if(dataRow == null) {
					dataRow = sheet.createRow(i + offsetYCounter);
				}
				if(columnSize > 1) {
					if (i != data.size() || !disableBottomRow){
						CellRangeAddress region = new CellRangeAddress(i + offsetYCounter, i + offsetYCounter, startOffsetX, startOffsetX + columnSize - 1);
						sheet.addMergedRegion(region);
						RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
					}
				}
				XSSFCell dataCell = dataRow.createCell(startOffsetX);
				// style the bottom rows
				if (data.size() == i) {
					if (!disableBottomRow){
						CellStyle bottomStyle = sheet.getWorkbook().createCellStyle();
						bottomStyle.cloneStyleFrom(columnStyle);
						bottomStyle.setDataFormat(columnStyle.getDataFormat());
						dataCell.setCellStyle(bottomStyle);
					}
				} else if(columnStyle != null) {
					if (reportData.getRowColorFunction() != null) {
						Object row = reportData.getElementList().get(i);
						String color = null;
						try {
							color = (String) row.getClass().getMethod(reportData.getRowColorFunction()).invoke(row);
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						} catch (InvocationTargetException e) {
							e.printStackTrace();
						} catch (NoSuchMethodException e) {
							e.printStackTrace();
						}
						XSSFCellStyle xssfCellStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
						xssfCellStyle.cloneStyleFrom(columnStyle);

//						DataFormatter formatter = new DataFormatter(Locale.forLanguageTag("tr-TR"));
//						formatter.addFormat("#.##0", new DecimalFormat("#.##0"));

						xssfCellStyle.setFillForegroundColor(new XSSFColor((java.awt.Color.decode(color))));
						xssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						dataCell.setCellStyle(xssfCellStyle);
//						formatter.formatCellValue(dataCell);
					} else {
						dataCell.setCellStyle(columnStyle);
					}
				}
				Locale turkey = new Locale("tr", "TR");
				NumberFormat turkishLirasFormat = NumberFormat.getCurrencyInstance(turkey);

				if (data.size() == i) {
					DecimalFormat df = new DecimalFormat();
					df.setMaximumFractionDigits(decimalPoint);
					boolean isNumber = false;
					if (!disableBottomRow){
						double sum = 0;
						int stringsCount = 0;
						for (T element:data) {
							if (element instanceof Float
									|| element instanceof Integer
									|| element instanceof Long
									|| element instanceof Double
									|| element instanceof BigDecimal
							) {
								sum += Double.parseDouble(element.toString());
							} else {
								stringsCount++;
							}
						}
						//build jitpack
						if (bottomCalculation == null || bottomCalculation.equals("string:"))
							dataCell.setCellValue("");
						else if (bottomCalculation != null && bottomCalculationText.equals("Tenzilat:")){
							Float d = Float.parseFloat(bottomValue);
							dataCell.setCellValue(bottomCalculationText+"\n"+ "%"+ df.format(d));
						}

						else if (bottomCalculation != null && bottomCalculationText.equals("TenzilatForGeneral:")){
							Float d = Float.parseFloat(bottomValue);
							dataCell.setCellValue("% "+ df.format(d));
						}

						else if (bottomCalculation != null && bottomCalculation.equals("avg"))
							dataCell.setCellValue(bottomCalculationText+"\n"+ sum / data.size());
						else if (bottomCalculation != null && bottomCalculation.equals("count"))
							dataCell.setCellValue(bottomCalculationText+"\n"+data.size());
						else if (bottomCalculation != null && bottomCalculation.equals("sum") && !bottomCalculation.equals("sumPercentage") && !bottomCalculation.equals("sumCount"))
							dataCell.setCellValue(!bottomCalculationText.equals("") ? bottomCalculationText+"\n"+turkishLirasFormat.format(sum)
												: turkishLirasFormat.format(sum));
						else if (bottomCalculation != null && bottomCalculation.equals("sumPercentage")){
							dataCell.setCellValue(sum);
							isNumber = true;
						}
						else if (bottomCalculation != null && bottomCalculation.equals("sumCount")){
							dataCell.setCellValue(sum);
							isNumber = true;
						}
						if (bottomCalculation != null && !bottomCalculation.equals("string:") && bottomCalculation.split(":")[0].equals("string") && stringsCount == data.size()) {
							dataCell.setCellValue(bottomCalculation.split(":")[1]);
						}
						offsetYCounter++;
					}
					if(bottomTitle != null){
						CellRangeAddress region = new CellRangeAddress(i + offsetYCounter, i + offsetYCounter, startOffsetX, startOffsetX + columnSize - 1);
						sheet.addMergedRegion(region);

						dataCell.setCellStyle(getBottomTitleCellStyle(sheet,12));
						dataCell.setCellValue(bottomTitle);
						bottomTitle = null;
						offsetYCounter++;
					}
					if(!isNumber && !dataCell.getStringCellValue().isEmpty()){
						reformatCell(dataCell, columnSize);
					}
				}
				else if(data.get(i) instanceof Float) {
					if(isDiscount != null && isDiscount==true && decimalPoint!= null && decimalPoint==0){
						Float d = Float.parseFloat(data.get(i).toString())*100;
						dataCell.setCellValue(Math.round(d)/100f);
					}else if(isDiscount != null && isDiscount==true && decimalPoint!=null && decimalPoint==1){
						dataCell.setCellValue(Float.parseFloat(data.get(i).toString()));
					}
					else{
						dataCell.setCellValue(Float.parseFloat(data.get(i).toString()));
					}
				}
				else if(data.get(i) instanceof Integer) {
					dataCell.setCellValue(Integer.parseInt(data.get(i).toString()));
				}
				else if (data.get(i) instanceof BigDecimal) {
					dataCell.setCellValue(((BigDecimal) data.get(i)).doubleValue());
				}
				else if (data.get(i) instanceof String) {
					if(cellContent.equals("link")){
						CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
						Hyperlink link = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
						if(data.get(i).toString().equals(" YILI TENZİLAT")){
							link.setAddress("' " + this.reportData.getYear() + " YILI TENZİLAT'!A1");
						}
						else if(data.get(i).toString().equals("  YILI TENZİLAT")){
							link.setAddress("' " + (this.reportData.getYear() - 1) + " YILI TENZİLAT'!A1");
						}
						else{
							link.setAddress("'" + data.get(i).toString() + "'!A1");
						}
						dataCell.setCellValue("Sayfaya git");
						dataCell.setHyperlink(link);
					}
					else{
						if(data.get(i).toString().equals(" YILI TENZİLAT")){
							dataCell.setCellValue(" " + this.reportData.getYear() + " YILI TENZİLAT");
						}
						else if(data.get(i).toString().equals("  YILI TENZİLAT")){
							dataCell.setCellValue(" " + (this.reportData.getYear() - 1) + " YILI TENZİLAT");
						}
						else{
							dataCell.setCellValue(data.get(i).toString());
						}
					}

					dataCell.setCellValue(dataCell.getStringCellValue().trim());

					if(!dataCell.getStringCellValue().isEmpty())
						reformatCell(dataCell, columnSize);;
				}
				else if (data.get(i) instanceof Long) {
					dataCell.setCellValue((Long) data.get(i));
				}
				else if (data.get(i) instanceof Double) {
					dataCell.setCellValue((Double) data.get(i));
				}
				else if (data.get(i) != null) {
					dataCell.setCellValue(data.get(i).toString());
				}
			}
			offsetYCounter += data.size();
		}

		public int getStartOffsetY() {
			return startOffsetY;
		}
	}

	private static class TableMapperExtended {
		private String header;
		private List<ColumnDefinition> columnDefinitionList;
		private ReportData reportData;
		private int startOffsetY;
		private int startOffsetX;

		private int offsetXCounter;

		public TableMapperExtended(String header, List<ColumnDefinition> columnDefinitionList, ReportData reportData) {
			this.header = header;
			this.columnDefinitionList = columnDefinitionList;
			this.reportData = reportData;
		}

		public void setStartOffsetY(int startOffsetY) {
			this.startOffsetY = startOffsetY;
		}

		public void setStartOffsetX(int startOffsetX) {
			this.startOffsetX = startOffsetX;
		}

		private Cell mergeCellAndSetBorder(XSSFSheet sheet, int rowStart, int rowEnd, int colStart, int colEnd){

			CellStyle borderStyle;
			if(rowStart == 1 && colStart == 1)
				borderStyle = getTitleHeaderStyle(sheet, reportData.titleFontSize);
			else{
				borderStyle = getTitleHeaderStyle(sheet, reportData.headerFontSize);
			}

			borderStyle.setBorderBottom(BorderStyle.THIN);
			borderStyle.setBorderLeft(BorderStyle.THIN);
			borderStyle.setBorderRight(BorderStyle.THIN);
			borderStyle.setBorderTop(BorderStyle.THIN);
			borderStyle.setAlignment(HorizontalAlignment.CENTER);

			for(int j = rowStart; j <= rowEnd; j++){
				Row row = sheet.getRow(j);
				if(row == null)
					row = sheet.createRow(j);

				for (int i = colStart; i <= colEnd; ++i) {
					Cell cell = row.createCell(i);
					cell.setCellStyle(borderStyle);
				}
			}
			sheet.addMergedRegion(new CellRangeAddress(rowStart, rowEnd, colStart, colEnd));
			return sheet.getRow(rowStart).getCell(colStart);
		}

		public void write(XSSFSheet sheet, ReportData reportData) {
			Row headerRow = sheet.getRow(startOffsetY);
			if(headerRow == null) {
				headerRow = sheet.createRow(startOffsetY);
			}
			int rowSize = 1;

			if(reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")
					|| reportData.reportType.equals("Ön Mali Kontrol İşlem Belgesi")
					|| reportData.reportType.substring(0,1).equals(" ")){

				Cell headerCell = mergeCellAndSetBorder(sheet,reportData.headerStartOffsetY, reportData.headerEndOffsetY, reportData.headerStartOffsetX, reportData.headerEndOffsetX);

				rowSize = reportData.headerEndOffsetY - reportData.headerStartOffsetY + 1;
				String title;
				if(reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
					title = reportData.year.toString()+" YILI" + "\n"
						+ header + "\n"
							+ (reportData.biddingDepartment.equals("Veri Girilmemiştir")? "": reportData.biddingDepartment.toUpperCase(Locale.ROOT)+", ")
						+ (reportData.biddingType.equals("Veri Girilmemiştir") ? "": reportData.biddingType.toUpperCase(Locale.ROOT)+", ")
						+ (reportData.biddingProcedure.equals("Veri Girilmemiştir") ? "": reportData.biddingProcedure.toUpperCase(Locale.ROOT));
				}else if(reportData.reportType.equals(" BÖLGEYE GÖRE DAĞILIM")

							|| reportData.reportType.equals(" İHALE TÜRÜNE GÖRE DAĞILIM")
							|| reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIM \n(YAPIM ve YAPIM(BAKIM) İHALELERİ)")
				){
					title = reportData.year.toString()+ " YILI ÖN MALİ KONTROLÜ YAPILAN İHALELER\n"+
								reportData.reportType;
				}else if(reportData.reportType.equals(" GENEL MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" BÖLGE MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" GENEL MÜDÜRLÜK & BÖLGE MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" MAL ALIM İŞİ İHALELERİ")
							|| reportData.reportType.equals(" YAPIM(BAKIM) İŞİ İHALELERİ")
							|| reportData.reportType.equals(" YAPIM İŞİ İHALELERİ")
							|| reportData.reportType.equals(" DANIŞMANLIK İŞİ İHALELERİ")
							|| reportData.reportType.equals(" HİZMET İŞİ İHALELERİ")
				){
					title = reportData.year.toString()+ " YILI ÖN MALİ KONTROL\n"+
							reportData.reportType;
				}
				else if(reportData.reportType.equals(" YAPIM İHALE USULE GÖRE DAĞILIMI")
						|| reportData.reportType.equals(" YAPIM İHALE USULE GÖRE TUTAR DAĞILIMI")
				){
					title = reportData.year.toString() + " YILI" + reportData.reportType;
				}else if(reportData.reportType.equals(" YAPIM İHALE USULE GÖRE TENZİLAT DAĞILIMI")){
					title = reportData.year.toString() + " YILI" + reportData.reportType;
				}
				else if(reportData.reportType.equals(" "+ reportData.year+ " YILI TENZİLAT")){
					title = reportData.reportType+"\n Yapım ve Yapım (Bakım) İhaleleri";
				}
				else if(reportData.reportType.equals(" TENZİLAT TABLO \n( SON 2 YIL )")){
					Integer previousYear = reportData.year-1;
					title = (previousYear)+"-"+(reportData.year)+" YILI" + reportData.reportType+"\n Yapım ve Yapım (Bakım) İhaleleri";
				}else if(reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIMI") ||
						reportData.reportType.equals(" İHALE USULÜNE GÖRE TUTAR DAĞILIMI") ||
						reportData.reportType.equals(" İHALE USULÜNE GÖRE TENZİLAT DAĞILIMI")){
					title = reportData.year.toString() + " YILI YAPIM İHALE TUTARININ\n"+reportData.reportType;
				}
				else if(reportData.reportType.equals(" İÇİNDEKİLER")){
					title = reportData.year + " YILI ÖN MALİ KONTROL RAPOR\n" + reportData.reportType;
				}
				else if(reportData.reportType.equals(" YILLARA GÖRE TENZİLAT BÖLGELER")){
					title ="YILLARA GÖRE TENZİLAT TABLOSU\nBÖLGELER";
				}
				else{
					title=header;
				}

				headerCell.setCellValue(title);

				offsetXCounter = startOffsetX;
				startOffsetY = reportData.headerEndOffsetY;

				if(reportData.reportType.equals(" YILLARA GÖRE ÖN MALİ KONTROL")){
					Cell subTitleRowCell1 = mergeCellAndSetBorder(sheet,startOffsetY+1, startOffsetY+1, reportData.headerStartOffsetX+2, reportData.yearCount+1);
					subTitleRowCell1.setCellValue("DOSYA SAYISI");

					Cell typeCell = mergeCellAndSetBorder(sheet,startOffsetY+1, startOffsetY+2, reportData.headerStartOffsetX, reportData.headerStartOffsetX+1);
					typeCell.setCellValue("İHALE TÜRÜ");

					Cell subTitleRowCell2 = mergeCellAndSetBorder(sheet, startOffsetY+1, startOffsetY+1, reportData.yearCount+2, reportData.yearCount*2+1);
					subTitleRowCell2.setCellValue("İHALE TUTARI (x1.000.000 TL)");
					startOffsetY+=1;
				}
			}
			else{
				Cell headerRowCell = headerRow.createCell(startOffsetX);
				headerRowCell.setCellStyle(getTitleHeaderStyle(sheet, reportData.titleFontSize));
				headerRowCell.setCellValue(header);

				offsetXCounter = startOffsetX;
			}

			double height;
			if (reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
				height = 17.0;

			}else if (reportData.reportType.substring(0,1).equals(" ")){
				height = 6.0;
			}
			else{
				height = 7.0;
			}

			int rowNum = headerRow.getRowNum();


			for(int i = 0; i < rowSize; i++){
				sheet.getRow(rowNum).setHeight((short) height);
				sheet.getRow(rowNum).setHeightInPoints((4* sheet.getRow(rowNum).getHeight()));
				rowNum++;
			}

			for (int i = 0; i < columnDefinitionList.size(); i++) {
				columnDefinitionList.get(i).write(sheet, startOffsetY + 1, offsetXCounter);
				offsetXCounter += columnDefinitionList.get(i).getColumnSize();
			}
			if (reportData.getTotalSumTitle() != null) {
				this.addTotalSumCell(sheet);
			}
		}

		private void addTotalSumCell(XSSFSheet sheet) {
			AtomicLong totalSum = new AtomicLong();
			reportData.getAddToTotalSumList().forEach((key) -> {
				String methodName = reportData.getColumnToMetadataMapping().get(key).getFunctionName();
				for (Object data: reportData.getElementList()) {
					try {
						totalSum.addAndGet((long)(double)data.getClass().getMethod(methodName).invoke(data));
					} catch (IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
						e.printStackTrace();
					}
				}
			});
			long summedValue = totalSum.get();
			int offsetY = startOffsetY + columnDefinitionList.size() + 1;
			Row totalRow = sheet.createRow(offsetY);
			CellRangeAddress region = new CellRangeAddress(offsetY, offsetY, startOffsetX, startOffsetX + 1);
			sheet.addMergedRegion(region);
			RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
			RegionUtil.setBorderTop(BorderStyle.MEDIUM, region, sheet);
			RegionUtil.setBorderLeft(BorderStyle.MEDIUM, region, sheet);
			RegionUtil.setBorderRight(BorderStyle.MEDIUM, region, sheet);
			Cell titleCell = totalRow.createCell(startOffsetX);
			titleCell.setCellValue(reportData.getTotalSumTitle());
			titleCell.setCellStyle(getHeaderRowStyle(sheet, reportData.headerFontSize));

			Cell sumCell = totalRow.createCell(startOffsetX + 2);
			sumCell.setCellValue(summedValue);
			sumCell.setCellStyle(getHeaderRowStyle(sheet, reportData.headerFontSize));
		}

		public void addChart(XSSFSheet sheet, List data, ChartProps chartProps, int chartOrder) {
			ChartDrawingService drawer = null;
			try {

				int defaultOffsetY = columnDefinitionList.get(0).offsetYCounter + 1;
				startOffsetY = defaultOffsetY + 21 * chartOrder;

				XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
				XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, startOffsetY, 7, startOffsetY + 20);
				anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
				XDDFChart chart = drawing.createChart(anchor);
				chart.setTitleText(chartProps.getTitle());
				chart.setTitleOverlay(false);

				drawer = new ChartDrawingService(chartProps.getTitle(), chartProps.getGroupLabel(), chartProps.getValueLabel(), chart);
				drawer.addData(
								data, chartProps.getGroupFunctionName(), chartProps.getValueFunctionName(),chartProps.getGroupLabel())
						.draw(chartProps.getType());

			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
