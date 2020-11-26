package com.paraboly.reportlib;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.Format;
import java.text.NumberFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Function;

import static com.paraboly.reportlib.utils.StyleUtils.*;

// build

public class GenericReports {
	@Data
	public static class ReportData {
		private List<?> elementList;
		private LinkedHashMap<String, ColumnMetadata> columnToMetadataMapping;
		private String reportType;
		private Integer year;
		private LinkedList<ChartProps> chartPropsLinkedList;
		private ChartProps chartProps;
		private LinkedList<String> addToTotalSumList;
		private String totalSumTitle;
		private Boolean disableBottomRow = false;
		private String rowColorFunction;
	}

	@Data
	public static class ColumnMetadata {
		private String functionName;
		private Function customFunction;
		private Integer columnSize = 1;
		private String bottomCalculation = "sum"; // potential values are sum, avg, or string:BOTTOM_NAME
		private String cellContent = "money"; // potential values are money, percentage, count, year
	}

	@Data
	public static class ChartProps {
		private String groupFunctionName;
		private String groupLabel;
		private String valueFunctionName;
		private String valueLabel;
		private String type;
		private String title;
		private String groupKey;
		private String valueKey;
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
				TableMapperExtended tableMapperExtended = getReportTable(reportData, sheet);
				tableMapperExtended.setStartOffsetX(2);
				tableMapperExtended.setStartOffsetY(2);
				tableMapperExtended.write(sheet);
				if (reportData.chartPropsLinkedList != null) {
					AtomicInteger i = new AtomicInteger(0);
					reportData.chartPropsLinkedList.forEach(chartProps -> {
						fillChartProps(chartProps, reportData.getColumnToMetadataMapping());
						tableMapperExtended.addChart(sheet, reportData.getElementList(), chartProps, i.getAndIncrement());
					});
				}
			}
			return wb;
		}

		private static ChartProps fillChartProps(ChartProps chartProps, LinkedHashMap<String, ColumnMetadata> columnMetadata) {
			chartProps.setGroupFunctionName(columnMetadata.get(chartProps.getGroupKey()).getFunctionName());
			chartProps.setGroupLabel(chartProps.getGroupKey());

			chartProps.setValueFunctionName(columnMetadata.get(chartProps.getValueKey()).getFunctionName());
			chartProps.setValueLabel(chartProps.getValueKey());
			return chartProps;
		}

		private static TableMapperExtended getReportTable(ReportData reportData, XSSFSheet sheet) {
			CellStyle dataStyle = getBorderedBoldCellStyle(sheet);
			CellStyle headerStyle = getHeaderRowStyle(sheet);

			CellStyle currencyStyle = sheet.getWorkbook().createCellStyle();
			currencyStyle.cloneStyleFrom(dataStyle);
			setCurrency(sheet, currencyStyle);

			CellStyle percentageStyle = sheet.getWorkbook().createCellStyle();
			percentageStyle.cloneStyleFrom(dataStyle);
			setPercentage(sheet, percentageStyle);

			CellStyle countStyle = sheet.getWorkbook().createCellStyle();
			countStyle.cloneStyleFrom(dataStyle);
			setCount(sheet, countStyle);

			CellStyle yearStyle = sheet.getWorkbook().createCellStyle();
			yearStyle.cloneStyleFrom(dataStyle);
			setYear(sheet, yearStyle);

			LinkedHashMap<String, ColumnDefinition> map = new LinkedHashMap<>();
			reportData.getColumnToMetadataMapping().forEach((columnName, columnMetadata) -> {
				CellStyle fieldStyle = null;
				switch (columnMetadata.getCellContent()) {
					case "money":
						fieldStyle = currencyStyle;
						break;
					case "percentage":
						fieldStyle = percentageStyle;
						break;
					case "count":
						fieldStyle = countStyle;
						break;
					case "year":
						fieldStyle = yearStyle;
						break;
				}
				map.put(columnName,
						new ColumnDefinition<String>(
								columnMetadata.getColumnSize(), columnName.toUpperCase(), fieldStyle, headerStyle,
								columnMetadata.getBottomCalculation(), reportData.getDisableBottomRow(), reportData));
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
		private Boolean disableBottomRow;
		private ReportData reportData;

		public ColumnDefinition(int columnSize, String column, CellStyle columnStyle, CellStyle headerStyle, String bottomCalculation, Boolean disableBottomRow, ReportData reportData) {
			this.columnSize = columnSize;
			this.column = column;
			this.columnStyle = columnStyle;
			this.headerStyle = headerStyle;
			this.bottomCalculation = bottomCalculation;
			this.disableBottomRow = disableBottomRow;
			this.reportData = reportData;
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

		public void write(Sheet sheet, int startOffsetY, int startOffsetX) {
			sheet.setDefaultColumnWidth(17);
			this.startOffsetX = startOffsetX;
			this.startOffsetY = startOffsetY;

			offsetYCounter = startOffsetY;
			Row columnHeaderRow = sheet.getRow(offsetYCounter);
			if(columnHeaderRow == null) {
				columnHeaderRow = sheet.createRow(offsetYCounter);
			}
			if(columnSize > 1) {
				CellRangeAddress region = new CellRangeAddress(offsetYCounter, offsetYCounter, startOffsetX, startOffsetX + columnSize - 1);
				sheet.addMergedRegion(region);
				RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
				RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
			}
			Cell cell = columnHeaderRow.createCell(startOffsetX);
			cell.setCellValue(column);
			if(headerStyle != null)
				cell.setCellStyle(headerStyle);
			offsetYCounter += 1;
			columnHeaderRow.setHeight((short) -1);

			for (int i = 0; i <= data.size(); i++) {
				Row dataRow = sheet.getRow(i + offsetYCounter);
				if(dataRow == null) {
					dataRow = sheet.createRow(i + offsetYCounter);
				}
				if(columnSize > 1) {
					CellRangeAddress region = new CellRangeAddress(i + offsetYCounter, i + offsetYCounter, startOffsetX, startOffsetX + columnSize - 1);
					sheet.addMergedRegion(region);
					RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
					RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
				}
				Cell dataCell = dataRow.createCell(startOffsetX);
				// style the bottom rows
				if (data.size() == i) {
					if (!disableBottomRow){
						CellStyle bottomStyle = sheet.getWorkbook().createCellStyle();
						bottomStyle.cloneStyleFrom(headerStyle);
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

						DataFormatter formatter = new DataFormatter(Locale.forLanguageTag("tr-TR"));
						formatter.addFormat("#.##0", new DecimalFormat("#.##0"));

						xssfCellStyle.setFillForegroundColor(new XSSFColor((java.awt.Color.decode(color))));
						xssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//						dataCell.setCellStyle(xssfCellStyle);
						formatter.formatCellValue(dataCell);
					} else {
						dataCell.setCellStyle(columnStyle);
					}
				}

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
						if (bottomCalculation == null || bottomCalculation.equals("sum"))
							dataCell.setCellValue(sum);
						else if (bottomCalculation != null && bottomCalculation.equals("avg"))
							dataCell.setCellValue(sum / data.size());
						if (bottomCalculation != null && bottomCalculation.split(":")[0].equals("string") && stringsCount == data.size()) {
							dataCell.setCellValue(bottomCalculation.split(":")[1]);
						}
					}
				}
				else if(data.get(i) instanceof Float) {
					dataCell.setCellValue(Float.parseFloat(data.get(i).toString()));
				}
				else if(data.get(i) instanceof Integer) {
					dataCell.setCellValue(Integer.parseInt(data.get(i).toString()));
				}
				else if (data.get(i) instanceof BigDecimal) {
					dataCell.setCellValue(((BigDecimal) data.get(i)).doubleValue());
				}
				else if (data.get(i) instanceof String) {
					dataCell.setCellValue(data.get(i).toString());
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

		public void write(Sheet sheet) {
			Row headerRow = sheet.getRow(startOffsetY);
			if(headerRow == null) {
				headerRow = sheet.createRow(startOffsetY);
			}

			Cell headerRowCell = headerRow.createCell(startOffsetX);
			headerRowCell.setCellStyle(getHeaderStyle(sheet));
			headerRowCell.setCellValue(header + " Tablosu");

			offsetXCounter = startOffsetX;
			for (int i = 0; i < columnDefinitionList.size(); i++) {
				columnDefinitionList.get(i).write(sheet, startOffsetY + 1, offsetXCounter);
				offsetXCounter += columnDefinitionList.get(i).getColumnSize();
			}
			if (reportData.getTotalSumTitle() != null) {
				this.addTotalSumCell(sheet);
			}
		}

		private void addTotalSumCell(Sheet sheet) {
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
			titleCell.setCellStyle(getHeaderRowStyle(sheet));

			Cell sumCell = totalRow.createCell(startOffsetX + 2);
			sumCell.setCellValue(summedValue);
			sumCell.setCellStyle(getHeaderRowStyle(sheet));
		}

		public void addChart(Sheet sheet, List data, ChartProps chartProps, int chartOrder) {
			ChartDrawingService drawer = null;
			Integer pictureIndex = null;
			try {
				drawer = new ChartDrawingService(chartProps.getTitle(), chartProps.getGroupLabel(), chartProps.getValueLabel());
				drawer.addData(
						data, chartProps.getGroupFunctionName(), chartProps.getValueFunctionName(),chartProps.getGroupLabel())
						.draw(chartProps.getType());
				InputStream imageStream = drawer.getInputStream();
				pictureIndex =
						sheet.getWorkbook().addPicture(IOUtils.toByteArray(imageStream), Workbook.PICTURE_TYPE_JPEG);

				XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
				CreationHelper helper = sheet.getWorkbook().getCreationHelper();
				ClientAnchor anchor = helper.createClientAnchor();
				int defaultOffsetY = (columnDefinitionList.get(0).getStartOffsetY() + columnDefinitionList.get(0).getData().size() + 3);
				startOffsetY = defaultOffsetY + 26 * chartOrder;
				startOffsetX = 2;
				anchor.setCol1( 0 );
				anchor.setRow1(startOffsetY); // same row is okay
				anchor.setRow2(startOffsetY);
				anchor.setCol2( 1 );
				anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
				Picture pict = drawing.createPicture(anchor, pictureIndex);
				pict.resize();
			} catch (NoSuchMethodException e) {
				e.printStackTrace();
			} catch (InvocationTargetException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
