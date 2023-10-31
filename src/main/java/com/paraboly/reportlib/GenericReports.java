package com.paraboly.reportlib;

import com.paraboly.reportlib.utils.StyleUtils;
import com.sun.corba.se.spi.orbutil.threadpool.Work;
import lombok.Builder;
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
//build

public class GenericReports {

	@Data
	public static class SheetData {
		private String sheetType;
		private int zoomLevel = 100;
		private List<ReportData> reportDataList;
		private int currentY;
	}
	@Data
	public static class ReportData {
		private List<?> elementList;
		private LinkedHashMap<String, ColumnMetadata> columnToMetadataMapping;
		private String reportType;
		private int fontSize=12;
		private int headerFontSize=12;
		private int titleFontSize=14;
		private List<Integer> yearList;
		private Integer headerStartOffsetX;
		private Integer headerEndOffsetX;
		private Integer headerStartOffsetY;
		private Integer headerEndOffsetY;
		private List<String> biddingDepartmentList;
		private List<String> biddingTypeList;
		private List<String> biddingProcedureList;
		private LinkedList<ChartProps> chartPropsLinkedList;
		private ChartProps chartProps;
		private LinkedList<String> addToTotalSumList;
		private String totalSumTitle;
		private Boolean disableBottomRow = false;
		private String rowColorFunction;
		private Integer yearCount;
		private Boolean mergeTwoRow = false;
		private SheetData sheetData;
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
		private boolean isReversed;
		private boolean isCustom;
		private String valueFormat;
		private String barDirection;
	}

	public static class Builder {
		private List<SheetData> sheetDataList;
		private String filename;
		private XSSFWorkbook wb;

		public Builder(String filename) {
			this.filename = filename;
			sheetDataList = new ArrayList<>();
			wb = new XSSFWorkbook();
		}

		public Builder addSheetData(SheetData sheetData) {
			sheetDataList.add(sheetData);
			return this;
		}

		public XSSFWorkbook create() {
			for (SheetData sheetData: sheetDataList){
				String sheetTitle = sheetData.getSheetType().equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER") ?
						" LİSTE" : sheetData.getSheetType();

				XSSFSheet sheet = wb.createSheet(sheetTitle);

				sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
				sheet.setFitToPage(true);
				sheet.getPrintSetup().setFitWidth((short)1);
				sheet.getPrintSetup().setFitHeight((short)1);
				sheet.setZoom(sheetData.zoomLevel);


				for(ReportData reportData : sheetData.getReportDataList()){
					reportData.setSheetData(sheetData);
					reportData.setHeaderStartOffsetY(sheetData.currentY + reportData.getHeaderStartOffsetY());
					reportData.setHeaderEndOffsetY(sheetData.currentY + reportData.getHeaderEndOffsetY());
					TableMapperExtended tableMapperExtended = getReportTable(reportData, sheet);
					if(sheetData.sheetType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")
							|| sheetData.sheetType.equals("CUMHURBAŞKANLIĞI")
							|| sheetData.sheetType.equals("BAKAN OLURLARI")
							|| sheetData.sheetType.equals(" İHALE USULÜNE GÖRE TUTAR DAĞILIMI")
							|| sheetData.sheetType.equals(" İHALE USULÜNE GÖRE TENZİLAT DAĞILIMI")
							|| sheetData.sheetType.equals(" İHALE USULÜNE GÖRE ORAN DAĞILIMI")
							|| sheetData.sheetType.equals(" GN. MD. BİLGİ NOTU")
							|| sheetData.sheetType.equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")
							|| (reportData.yearList != null && !reportData.yearList.isEmpty() &&
								sheetData.sheetType.equals(" "+ reportData.yearList.get(0)+ " YILI TENZİLAT"))
							|| sheetData.sheetType.equals(" TENZİLAT TABLO \n( SON 2 YIL )")
					)
					{
						tableMapperExtended.setStartOffsetX(reportData.headerStartOffsetX);
						tableMapperExtended.setStartOffsetY(reportData.headerStartOffsetY);
					}
					tableMapperExtended.write(sheet, reportData);
					if (reportData.chartPropsLinkedList != null) {
						AtomicInteger i = new AtomicInteger(0);
						reportData.chartPropsLinkedList.forEach(chartProps -> {
							if(!chartProps.isReversed())
								fillChartProps(chartProps, reportData.getColumnToMetadataMapping());
							else
								fillChartPropsReversed(chartProps, reportData.getColumnToMetadataMapping());
							tableMapperExtended.addChart(sheet, reportData.getElementList(), chartProps, i.getAndIncrement(), reportData.chartPropsLinkedList.size());
						});
					}
					sheetData.currentY = tableMapperExtended.columnDefinitionList.get(0).offsetYCounter + 1;
				}
			}
			return wb;
		}

		private ChartProps fillChartProps(ChartProps chartProps, LinkedHashMap<String, ColumnMetadata> columnMetadata) {
			chartProps.setGroupFunctionName(columnMetadata.get(chartProps.getGroupKey()).getFunctionName());
			chartProps.setGroupLabel(chartProps.getGroupKey());

			chartProps.setValueFunctionName(new String[chartProps.valueKey.length]);

			for(int t = 0; t < chartProps.valueKey.length; t++){
				chartProps.getValueFunctionName()[t] = columnMetadata.get(chartProps.getValueKey()[t]).getFunctionName();
			}

			chartProps.setValueLabel(chartProps.getValueKey());
			return chartProps;
		}
		private ChartProps fillChartPropsReversed(ChartProps chartProps, LinkedHashMap<String, ColumnMetadata> columnMetadata) {
			if(chartProps.getGroupLabel() == null)
				chartProps.setGroupLabel(chartProps.getGroupKey());

			if(chartProps.getValueLabel() == null)
				chartProps.setValueLabel(chartProps.getValueKey());
			return chartProps;
		}
		private CellStyle getCellStyle(XSSFSheet sheet, String type, ColumnMetadata columnMetadata, int size){
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


		private TableMapperExtended getReportTable(ReportData reportData, XSSFSheet sheet) {
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

		private Object invokeCustomMethod(Object data, Function function) {
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

		public void reformatCell(XSSFCell dataCell, int columnSize){
			String fontName = dataCell.getCellStyle().getFont().getFontName();
			int fontSize = dataCell.getCellStyle().getFont().getFontHeightInPoints();
			final DataFormatter dataFormatter = new DataFormatter();
			final FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator(dataCell.getSheet().getWorkbook());
			objFormulaEvaluator.evaluate(dataCell);
			final String value = dataFormatter.formatCellValue(dataCell, objFormulaEvaluator);

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
					dataCell.getRow().setHeightInPoints(length * 2f);
			}
			else if(textheight * 1.5f > dataCell.getRow().getHeightInPoints()) {
				dataCell.getRow().setHeightInPoints(textheight * 2f);
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
			if (this.reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER") || this.reportData.reportType.equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")){
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
			Cell cell = columnHeaderRow.createCell(startOffsetX);
			if(!(isMerged && column.equals("İHALE TÜRÜ"))) { // control
				cell.setCellValue(column);
				if (headerStyle != null)
					cell.setCellStyle(headerStyle);
			}
			if(this.reportData.reportType.equals("CUMHURBAŞKANLIĞI") || this.reportData.reportType.equals("BAKAN OLURLARI")
					|| this.reportData.sheetData.sheetType.equals("GN. MD. BİLGİ NOTU")){
				XSSFCellStyle xssfCellStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
				xssfCellStyle.cloneStyleFrom(headerStyle);
				xssfCellStyle.setFillForegroundColor(new XSSFColor((java.awt.Color.decode("#5b9bd5"))));
				xssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				xssfCellStyle.setFont(getHeaderFont(sheet,12));
				cell.setCellStyle(xssfCellStyle);
			}


			for (int i = 0; i <= data.size(); i++) {
				XSSFRow dataRow = sheet.getRow(i + offsetYCounter);
				if(dataRow == null) {
					dataRow = sheet.createRow(i + offsetYCounter);
				}
				boolean notDeflator = true;
				if(i < data.size() && reportData.reportType.equals(" İÇİNDEKİLER")){
					Object row = reportData.getElementList().get(i);
					String reportName = null;
					try{
						reportName = (String)(row.getClass().getMethod(reportData.getColumnToMetadataMapping().get("Rapor Adı").getFunctionName()).invoke(row)); // only for those have function name
					}
					catch(NoSuchMethodException | IllegalAccessException | InvocationTargetException | ClassCastException e){
						e.printStackTrace();
					}
					if(reportName != null && reportName.equals(" YILLARA GÖRE ÖN MALİ K.(GÜNCEL)")){
						CellRangeAddress region = new CellRangeAddress(i + offsetYCounter, i + offsetYCounter + 1, startOffsetX, startOffsetX + columnSize - 1);
						sheet.addMergedRegion(region);
						RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
						RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
						offsetYCounter++;
						notDeflator = false;
					}
				}
				if(columnSize > 1 && notDeflator) {
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
					if (!disableBottomRow) {
						CellStyle bottomStyle = sheet.getWorkbook().createCellStyle();
						bottomStyle.cloneStyleFrom(columnStyle);
						bottomStyle.setDataFormat(columnStyle.getDataFormat());
						dataCell.setCellStyle(bottomStyle);
						if (bottomCalculation != null &&
								bottomCalculationText != null &&
								!bottomCalculationText.isEmpty()) {

							DataFormat format = sheet.getWorkbook().createDataFormat();
							String formatAsString = columnStyle.getDataFormatString();
							if(bottomCalculation.equals("count")){
								formatAsString = "#,##0";
							}
							bottomStyle.setDataFormat(format.getFormat("\"" + bottomCalculationText +
									"\n" + "\"" + formatAsString));

						}
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
						if (bottomCalculation == null || bottomCalculation.equals("string:")) {
							dataCell.setCellValue("");
						}
						else if (bottomValue != null && !bottomValue.isEmpty()) {
							if(bottomCalculationText.equals("Toplam Yaklaşık Maliyet:") ||
								bottomCalculationText.equals("Toplam İhale Bedeli:")) {

								dataCell.setCellValue(Double.parseDouble(bottomValue));
							}
							else{
								float d = Float.parseFloat(bottomValue);
								dataCell.setCellValue(d);
							}

						}
						else if (bottomCalculation.equals("count")) {
							dataCell.setCellValue(data.size());
						}
						else if (bottomCalculation.equals("sum") ||
								bottomCalculation.equals("sumPercentage") ||
								bottomCalculation.equals("sumCount")) {
							dataCell.setCellValue(sum);
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

					reformatCell(dataCell, columnSize);

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
							link.setAddress("' " + this.reportData.getYearList().get(0) + " YILI TENZİLAT'!A1");
						}
						else if(data.get(i).toString().equals("  YILI TENZİLAT")){
							link.setAddress("' " + (this.reportData.getYearList().get(0) - 1) + " YILI TENZİLAT'!A1");
						}
						else if(data.get(i).toString().equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
							link.setAddress("' LİSTE'!A1");
						}else if(data.get(i).toString().equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")){
							link.setAddress("'SÖZLEŞME & ÖN MALİ KONTROL KARŞ'!A1");
						}
						else{
							link.setAddress("'" + data.get(i).toString() + "'!A1");
						}
						dataCell.setCellValue("Sayfaya git");
						dataCell.setHyperlink(link);
					}
					else{
						if(data.get(i).equals(" YILLARA GÖRE ÖN MALİ K.(GÜNCEL)")){
							assert reportData.yearList != null;
							int currentYear = reportData.yearList.get(0);
							int beginningYear = currentYear - reportData.yearCount + 1;
							dataCell.setCellValue(" YILLARA GÖRE ÖN MALİ KONTROL\n" + "( " + beginningYear + "-" + currentYear
									+ " YILLARI, " +  Calendar.getInstance().get(Calendar.YEAR) + " YILI FİYATLARIYLA )");
						}
						else if(data.get(i).toString().equals(" YILI TENZİLAT")){
							dataCell.setCellValue(" " + this.reportData.getYearList().get(0) + " YILI TENZİLAT");
						}
						else if(data.get(i).toString().equals("  YILI TENZİLAT")){
							dataCell.setCellValue(" " + (this.reportData.getYearList().get(0) - 1) + " YILI TENZİLAT");
						}
						else if(data.get(i).toString().equals(" DURUM")){
							dataCell.setCellValue(" ÖN MALİ KONTROL" + data.get(i).toString());
						}
						else if(data.get(i).toString().equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
							dataCell.setCellValue("ÖN MALİ KONTROL LİSTE");
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
			Row headerRow = sheet.getRow(reportData.getHeaderStartOffsetY()); // check here for other reports
			if(headerRow == null) {
				headerRow = sheet.createRow(reportData.getHeaderStartOffsetY()); // check here for other reports
			}
			int rowSize = 1;

			if(reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")
					|| reportData.reportType.equals("Ön Mali Kontrol İşlem Belgesi")
					|| reportData.reportType.substring(0,1).equals(" ")
					|| reportData.reportType.equals("CUMHURBAŞKANLIĞI")
					|| reportData.reportType.equals("BAKAN OLURLARI")
					|| reportData.sheetData.sheetType.equals("GN. MD. BİLGİ NOTU")
					|| reportData.sheetData.sheetType.equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")
			){

				Cell headerCell = mergeCellAndSetBorder(sheet,reportData.headerStartOffsetY, reportData.headerEndOffsetY, reportData.headerStartOffsetX, reportData.headerEndOffsetX);

				rowSize = reportData.headerEndOffsetY - reportData.headerStartOffsetY + 1;
				String title;
				if(reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER")){
					StringBuilder years = new StringBuilder();
					if(reportData.yearList != null) {
						for (Integer year : reportData.yearList) {
							years.append(year);
							years.append("-");
						}
						if(years.length() > 0)
							years.deleteCharAt(years.length() - 1);
					}

					StringBuilder biddingDepartments = new StringBuilder();
					if(reportData.biddingDepartmentList != null){
						for(String department : reportData.biddingDepartmentList){
							if(!department.equals("Veri Girilmemiştir")){
								biddingDepartments.append(department.toUpperCase(Locale.ROOT));
								biddingDepartments.append("-");
							}
						}
						if(biddingDepartments.length() > 0)
							biddingDepartments.deleteCharAt(biddingDepartments.length() - 1);
					}


					StringBuilder biddingTypes = new StringBuilder();
					if(reportData.biddingTypeList != null){
						for(String biddingType : reportData.biddingTypeList){
							if(!biddingType.equals("Veri Girilmemiştir")){
								biddingTypes.append(biddingType.toUpperCase(Locale.ROOT));
								biddingTypes.append("-");
							}
						}
						if(biddingTypes.length() > 0)
							biddingTypes.deleteCharAt(biddingTypes.length() - 1);
					}


					StringBuilder biddingProcedures = new StringBuilder();
					if(reportData.biddingProcedureList != null){
						for(String biddingProcedure : reportData.biddingProcedureList){
							if(!biddingProcedure.equals("Veri Girilmemiştir")){
								biddingProcedures.append(biddingProcedure.toUpperCase(Locale.ROOT));
								biddingProcedures.append("-");
							}
						}
						if(biddingProcedures.length() > 0)
							biddingProcedures.deleteCharAt(biddingProcedures.length() - 1);
					}
					assert reportData.yearList != null;
					title = (reportData.yearList.size() == 1 ? (years + " YILI") : (years + " YILLARI")) + "\n"
						+ header + "\n"
							+ (reportData.biddingDepartmentList == null || reportData.biddingDepartmentList.size() == 0 ? "" :  (biddingDepartments + ", "))
						+ (reportData.biddingTypeList == null || reportData.biddingTypeList.size() == 0 ? "" : (biddingTypes + ", "))
						+ (reportData.biddingProcedureList == null || reportData.biddingProcedureList.size() == 0 ? "" : (biddingProcedures));
				}else if(reportData.reportType.equals(" BÖLGEYE GÖRE DAĞILIM")

							|| reportData.reportType.equals(" İHALE TÜRÜNE GÖRE DAĞILIM")
							|| reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIM \n(YAPIM ve YAPIM(BAKIM) İHALELERİ)")
				){
					title = reportData.yearList.get(0).toString()+ " YILI ÖN MALİ KONTROLÜ YAPILAN İHALELER\n"+
								reportData.reportType;
				}
				else if(reportData.reportType.equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")){
					title = reportData.yearList.get(0).toString() + " YILI\n" + reportData.reportType;
				}
				else if(reportData.reportType.equals(" DURUM")){
					title = reportData.yearList.get(0).toString()+ " YILI ÖN MALİ KONTROL\n"+
							reportData.reportType;
				}
				else if(reportData.reportType.equals(" GENEL MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" BÖLGE MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" GENEL MÜDÜRLÜK & BÖLGE MÜDÜRLÜK İHALELERİ")
							|| reportData.reportType.equals(" MAL ALIM İŞİ İHALELERİ")
							|| reportData.reportType.equals(" YAPIM(BAKIM) İŞİ İHALELERİ")
							|| reportData.reportType.equals(" YAPIM İŞİ İHALELERİ")
							|| reportData.reportType.equals(" DANIŞMANLIK İŞİ İHALELERİ")
							|| reportData.reportType.equals(" HİZMET İŞİ İHALELERİ")
				){
					title = reportData.yearList.get(0).toString()+ " YILI ÖN MALİ KONTROL\n"+
							reportData.reportType;
				}
				else if(reportData.reportType.equals(" YAPIM İHALE USULE GÖRE DAĞILIMI")
						|| reportData.reportType.equals(" YAPIM İHALE USULE GÖRE TUTAR DAĞILIMI")
				){
					title = reportData.yearList.get(0).toString() + " YILI" + reportData.reportType;
				}else if(reportData.reportType.equals(" YAPIM İHALE USULE GÖRE TENZİLAT DAĞILIMI")){
					title = reportData.yearList.get(0).toString() + " YILI" + reportData.reportType;
				}
				else if(reportData.yearList != null && reportData.reportType.equals(" "+ reportData.yearList.get(0)+ " YILI TENZİLAT")){
					title = reportData.reportType+"\n Yapım ve Yapım (Bakım) İhaleleri";
				}
				else if(reportData.reportType.equals(" TENZİLAT TABLO \n( SON 2 YIL )")){
					Integer previousYear = reportData.yearList.get(0)-1;
					title = (previousYear)+"-"+(reportData.yearList.get(0))+" YILI" + reportData.reportType+"\n Yapım ve Yapım (Bakım) İhaleleri";
				}else if(reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIMI") ||
						reportData.reportType.equals(" İHALE USULÜNE GÖRE TUTAR DAĞILIMI") ||
						reportData.reportType.equals(" İHALE USULÜNE GÖRE ORAN DAĞILIMI") ||
						reportData.reportType.equals(" İHALE USULÜNE GÖRE TENZİLAT DAĞILIMI")){
					title = reportData.yearList.get(0).toString() + " YILI YAPIM İHALE TUTARININ\n"+reportData.reportType;
				}
				else if(reportData.reportType.equals(" İÇİNDEKİLER")){
					title = reportData.yearList.get(0) + " YILI ÖN MALİ KONTROL RAPOR\n" + reportData.reportType;
				}
				else if(reportData.reportType.equals(" YILLARA GÖRE TENZİLAT BÖLGELER")){
					title ="YILLARA GÖRE TENZİLAT TABLOSU\nBÖLGELER";
				}
				else if(reportData.reportType.equals(" YILLARA GÖRE ÖN MALİ K.(GÜNCEL)")){
					assert reportData.yearList != null;
					int currentYear = reportData.yearList.get(0);
					int beginningYear = currentYear - reportData.yearCount + 1;
					title ="YILLARA GÖRE ÖN MALİ KONTROL\n" + "( " + beginningYear + "-" + currentYear + " YILLARI, "
							+ Calendar.getInstance().get(Calendar.YEAR) + " YILI FİYATLARIYLA )";
				}
				else if(reportData.reportType.equals("CUMHURBAŞKANLIĞI")){
					title = "ULAŞTIRMA ve ALTYAPI BAKANLIĞI\n(CUMHURBAŞKANLIĞINDA BEKLEYEN)\nYATIRIM PROGRAMI REVİZYON TALEPLERİ TAKİP TABLOSU";
				}
				else if(reportData.reportType.equals("BAKAN OLURLARI")){
					title = "ULAŞTIRMA ve ALTYAPI BAKANLIĞINA\nBAKAN OLURU ALINMASI İÇİN İLETİLEN TALEPLER";
				}
				else{
					title=header;
				}

				headerCell.setCellValue(title);

				offsetXCounter = startOffsetX;
				startOffsetY = reportData.headerEndOffsetY;

				if(reportData.reportType.equals(" YILLARA GÖRE ÖN MALİ KONTROL") || reportData.reportType.equals(" YILLARA GÖRE ÖN MALİ K.(GÜNCEL)")){
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
			if (reportData.reportType.equals("ÖN MALİ KONTROLÜ YAPILAN İHALELER") || reportData.reportType.equals("SÖZLEŞME & ÖN MALİ KONTROL KARŞILAŞTIRMA RAPORU")){
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

		public void addChart(XSSFSheet sheet, List data, ChartProps chartProps, int chartOrder, int size) {
			if(data == null || data.size() == 0) return;
			ChartDrawingService drawer = null;
			try {

				int defaultOffsetY = columnDefinitionList.get(0).offsetYCounter + 1;
				startOffsetY = defaultOffsetY + 21 * chartOrder;
				int col2 = 7;
				int col1 = 0;
				if(reportData.reportType.equals(" İHALE TÜRÜNE GÖRE DAĞILIM")
						|| reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIM \n(YAPIM ve YAPIM(BAKIM) İHALELERİ)")
						|| reportData.reportType.equals(" GENEL MÜDÜRLÜK & BÖLGE MÜDÜRLÜK İHALELERİ")
						|| reportData.reportType.equals(" BÖLGE MÜDÜRLÜK İHALELERİ")
						|| reportData.reportType.equals(" GENEL MÜDÜRLÜK İHALELERİ")
						|| reportData.reportType.equals(" DURUM")
				){
					startOffsetY = columnDefinitionList.get(0).offsetYCounter + 1;
					col1 = chartOrder * (offsetXCounter / size);
					col2 = col1 + (offsetXCounter / size);
				}

				else if(reportData.yearList != null && reportData.reportType.equals(" "+ reportData.yearList.get(0)+ " YILI TENZİLAT")
						|| reportData.yearList != null && reportData.reportType.equals(" "+ (reportData.yearList.get(0) - 1) + " YILI TENZİLAT")
						|| reportData.reportType.equals(" TENZİLAT TABLO \n( SON 2 YIL )")){
					col2 = 8;
				}
				else if(reportData.reportType.equals(" BÖLGEYE GÖRE DAĞILIM")
						|| reportData.reportType.equals(" MAL ALIM İŞİ İHALELERİ")
						|| reportData.reportType.equals(" YAPIM(BAKIM) İŞİ İHALELERİ")
						|| reportData.reportType.equals(" YAPIM İŞİ İHALELERİ")
						|| reportData.reportType.equals(" DANIŞMANLIK İŞİ İHALELERİ")
						|| reportData.reportType.equals(" HİZMET İŞİ İHALELERİ")
						|| reportData.reportType.equals(" İHALE USULÜNE GÖRE ORAN DAĞILIMI")
						|| reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIMI")){
					col2 = offsetXCounter;
				}
				XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
				int chartHeight = 20;
				if(reportData.reportType.equals(" İHALE USULÜNE GÖRE DAĞILIMI") || reportData.reportType.equals(" İHALE USULÜNE GÖRE ORAN DAĞILIMI")){
					chartHeight = 45;
				}
				XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, col1, startOffsetY, col2, startOffsetY + chartHeight);
				anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
				XDDFChart chart = drawing.createChart(anchor);
				chart.setTitleText(chartProps.getTitle());
				chart.setTitleOverlay(false);

				drawer = new ChartDrawingService(chartProps.getTitle(), chartProps.getGroupLabel(), chartProps.getValueLabel(), chart, chartProps.getValueFormat(), chartProps.getBarDirection());

				if(chartProps.isCustom() && chartProps.isReversed()){
					drawer.addDataCustom(
									data, chartProps.getValueKey(), reportData.getColumnToMetadataMapping())
							.draw(chartProps.getType());
				}
				else if(chartProps.isReversed()){
					drawer.addDataReversed(
									data, chartProps.getValueKey(), reportData.getColumnToMetadataMapping())
							.draw(chartProps.getType());
				}
				else{
					drawer.addData(
									data, chartProps.getGroupFunctionName(), chartProps.getValueFunctionName(),chartProps.getGroupLabel(), chartProps.getValueKey(), reportData.getColumnToMetadataMapping())
							.draw(chartProps.getType());
				}


			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
