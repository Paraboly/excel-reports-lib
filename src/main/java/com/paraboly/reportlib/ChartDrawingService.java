package com.paraboly.reportlib;

import lombok.SneakyThrows;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.labels.PieSectionLabelGenerator;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.DefaultDrawingSupplier;
import org.jfree.chart.plot.PiePlot3D;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.util.DefaultShadowGenerator;
import org.jfree.chart.util.TableOrder;
import org.jfree.data.category.CategoryToPieDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.main.*;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class ChartDrawingService {

	public class ChartDrawingSupplier extends DefaultDrawingSupplier {

		public Paint[] paintSequence;
		public int paintIndex;
		public int fillPaintIndex;

		{
			paintSequence =  new Paint[] {
					new Color(78, 127, 187),
					new Color(189,79, 76),
					new Color(153,185,88),
					new Color(126,99,159),
					new Color(230,218,119),
					new Color(60,201,178),
					new Color(214,158,36),
					new Color(240,117,225)
			};
		}

		@Override
		public Paint getNextPaint() {
			Paint result
					= paintSequence[paintIndex % paintSequence.length];
			paintIndex++;
			return result;
		}


		@Override
		public Paint getNextFillPaint() {
			Paint result
					= paintSequence[fillPaintIndex % paintSequence.length];
			fillPaintIndex++;
			return result;
		}
	}
	private String valueFormat;
	private DefaultCategoryDataset dataset;
	private String[] categories;
	private Double[][] values;
	private XDDFChart XDDFchart;
	private String title, categoryLabel;
	private String[]  valueLabel;
	private String barDirection;
	private JFreeChart chart;
	private byte[][] colors =  new byte[][] {
				new byte[]{(byte)78, (byte)127, (byte)187},
				new byte[]{(byte)189, (byte)79, (byte)76},
				new byte[]{(byte)153, (byte)185, (byte)88},
				new byte[]{(byte)126, (byte)99, (byte)159},
				new byte[]{(byte)230, (byte)218, (byte)119},
				new byte[]{(byte)60, (byte)201, (byte)178},
				new byte[]{(byte)214, (byte)158, (byte)36},
				new byte[]{(byte)240, (byte)117, (byte)225}
	};

	public Integer getHeight() {
		return height;
	}

	public void setHeight(Integer height) {
		this.height = height;
	}

	public Integer getWidth() {
		return width;
	}

	public void setWidth(Integer width) {
		this.width = width;
	}

	private Integer height = 480;
	private Integer width = 640;

	public ChartDrawingService(String title, String categoryLabel, String[] valueLabel, XDDFChart XDDFchart, String valueFormat, String barDirection){
		dataset = new DefaultCategoryDataset();
		this.title = title;
		this.categoryLabel = categoryLabel;
		this.valueLabel = valueLabel;
		this.XDDFchart = XDDFchart;
		this.valueFormat = valueFormat;
		this.barDirection = barDirection;
	}

	public ChartDrawingService addData(List<?> dataList, String categoryMethod, String[] valueMethod, String groupName, String[] valueKey, LinkedHashMap<String, GenericReports.ColumnMetadata> colData) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
		categories = new String[dataList.size()];
		values = new Double[valueMethod.length][dataList.size()];

		for(int t = 0; t < valueMethod.length; t++){

			int i = 0, j = 0;

			for (Object data: dataList) {

				String category = null;
				try{
					category = (String) data.getClass().getMethod(categoryMethod).invoke(data);
				}
				catch(Exception e){
					category = (colData.get(groupName).getCustomFunction().apply(data).toString());
				}


				Double value = null;
				try{
					value = Double.valueOf(data.getClass().getMethod(valueMethod[t]).invoke(data).toString());
				}
				catch (Exception e) {
					value = Double.valueOf(colData.get(valueKey[t]).getCustomFunction().apply(data).toString());
				}

				//dataset.addValue(value, groupName, category);
				categories[i++] = category;
				values[t][j++] = value;
			}
		}

		return this;
	}
	public ChartDrawingService addDataReversed(List<?> dataList, String[] valueKey, LinkedHashMap<String, GenericReports.ColumnMetadata> colData) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
		categories = valueKey;
		values = new Double[dataList.size()][valueKey.length];
		int i = 0;
		for (Object data: dataList) {
			for(int t = 0; t < valueKey.length; t++) {
				Double value = Double.valueOf(colData.get(valueKey[t]).getCustomFunction().apply(data).toString());
				values[i][t] = value;
			}
			i++;
		}
		return this;
	}
	public ChartDrawingService addDataCustom(List<?> dataList, String[] valueKey, LinkedHashMap<String, GenericReports.ColumnMetadata> colData)throws NoSuchMethodException, InvocationTargetException, IllegalAccessException{

		categories = new String[valueKey.length / 2];
		for(int i = 0; i < categories.length; i++){
			categories[i] = valueKey[i];
		}
		Double[] sumCountList = new Double[valueKey.length / 2];
		Double[] sumMoneyList = new Double[valueKey.length / 2];
		for(int t = 0; t < valueKey.length / 2; t++) {
			for (Object data: dataList) {
				Double value = Double.valueOf(colData.get(valueKey[t]).getCustomFunction().apply(data).toString());
				sumCountList[t] = sumCountList[t] == null ? value : sumCountList[t] + value;
			}
		}
		for(int t = valueKey.length / 2; t < valueKey.length; t++) {
			for (Object data: dataList) {
				Double value = Double.valueOf(colData.get(valueKey[t]).getCustomFunction().apply(data).toString());
				sumMoneyList[t - valueKey.length / 2] = sumMoneyList[t - valueKey.length / 2] == null ? value : sumMoneyList[t - valueKey.length / 2] + value;
			}
		}
		values = new Double[2][valueKey.length / 2];
		values[0] = sumCountList;
		values[1] = sumMoneyList;
		return this;
	}

	@SneakyThrows
	public void draw(String chartType) {
		switch (chartType) {
			case "bar":
				drawBarChartWithXDDF();
				break;
			case "line":
				drawLineChartWithXDDF();
				break;
			case "combinedBar":
				drawCombinedBarChartWithXDDF();
				break;
			case "combinedBarForDiscount":
				drawCombinedBarChartForDiscountWithXDDF();
				break;
			case "pie":
				drawPieChartWithXDDF();
				break;
			case "stackedBarChart":
				drawStackedBarChartWithXDDF();
				break;
			case "combinedStackedBarChart":
				drawCombinedStackedBarChartWithXDDF();
				break;
			case "combinedBarForCustom":
				drawCustomChartWithXDDF();
				break;
			case "combinedBarAndLine":
				drawCombinedBarAndLineChartWithXDDF();
			default:
				break;
		}
	}

	private JFreeChart drawBarChart() {
		JFreeChart barChartObject = ChartFactory.createBarChart(
				title,categoryLabel,
				valueLabel[0],
				this.dataset, PlotOrientation.VERTICAL,
				true,true,false);
		barChartObject.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions.UP_45);
		barChartObject.getCategoryPlot().getDomainAxis().setMaximumCategoryLabelLines(2);
		barChartObject.getPlot().setBackgroundPaint(Color.white);
		barChartObject.getPlot().setDrawingSupplier(new ChartDrawingSupplier());
		return barChartObject;
	}

	private JFreeChart drawPieChartWithJFree() {
		CategoryToPieDataset pieDataset = new CategoryToPieDataset(dataset, TableOrder.BY_ROW, 0);
		JFreeChart pieChartObject = ChartFactory.createPieChart3D(title, pieDataset,true,true,false);
		final PieSectionLabelGenerator labelGenerator = new StandardPieSectionLabelGenerator("{0} ({2})");
		final PiePlot3D plot = (PiePlot3D) pieChartObject.getPlot();
		final ChartPanel chartPanel = new ChartPanel(pieChartObject);
		chartPanel.setPreferredSize(new java.awt.Dimension(500, 100));
		plot.setLabelGenerator(labelGenerator);
		plot.setBackgroundPaint(Color.white);
		plot.setForegroundAlpha(0.8f);
		plot.setShadowGenerator(new DefaultShadowGenerator(5, Color.decode("#c4c4c4"),0.5f, 4,Math.PI/11));
		plot.setDrawingSupplier(new ChartDrawingSupplier());
		plot.setLabelBackgroundPaint(Color.white);

		return pieChartObject;
	}
	private void drawPieChartWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.BOTTOM);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(values[0]);

		XDDFChartData chartData = XDDFchart.createData(ChartTypes.PIE3D, null, null);

		chartData.setVaryColors(true);
		XDDFChartData.Series series = chartData.addSeries(cat, val);
		XDDFchart.plot(chartData);

		if (!XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).isSetDLbls())
			XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).addNewDLbls();
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowLegendKey().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowPercent().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowLeaderLines().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowVal().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowCatName().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowSerName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).getDLbls()
				.addNewShowBubbleSize().setVal(false);

		/*CTShapeProperties shapeProperties = XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0)
				.getSerArray(0).addNewSpPr();

		shapeProperties.addNewLn().addNewSolidFill()
				.addNewSrgbClr().setVal(new byte[]{(byte)255,(byte)255,(byte)255});*/

		int pointCount = series.getCategoryData().getPointCount();
		for (int p = 0; p < pointCount; p++) {

			XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0).getSerArray(0).addNewDPt()
					.addNewIdx().setVal(p);

			CTShapeProperties shapeProperties = XDDFchart.getCTChart().getPlotArea().getPie3DChartArray(0)
					.getSerArray(0).getDPtArray(p).addNewSpPr();

			shapeProperties.addNewSolidFill().addNewSrgbClr().setVal(colors[p % colors.length]);
			shapeProperties.addNewLn().addNewSolidFill().addNewSrgbClr()
					.setVal(new byte[]{(byte)255,(byte)255,(byte)255});
		}

		series.setFillProperties(new XDDFSolidFillProperties());

		if (XDDFchart.getCTChart().getAutoTitleDeleted() == null) XDDFchart.getCTChart().addNewAutoTitleDeleted();
		XDDFchart.getCTChart().getAutoTitleDeleted().setVal(false);

		Integer angle = 35;
		XDDFchart.getOrAddView3D().setXRotationAngle(angle.byteValue());
	}

	private void drawBarChartWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.BOTTOM);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);
		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle(valueLabel[0]);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(values[0]);

		XDDFChartData chartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		chartData.setVaryColors(true);
		XDDFChartData.Series series = chartData.addSeries(cat, val);
		series.setTitle(categoryLabel, null);
		if(this.valueFormat != null){
			XDDFchart.getCTChart().getPlotArea().getValAxArray(0).addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getValAxArray(0).getNumFmt().setFormatCode(this.valueFormat);

			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);

			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().getNumFmt().setFormatCode(this.valueFormat);
		}
		XDDFchart.plot(chartData);

		XDDFBarChartData bar = (XDDFBarChartData) chartData;
		if(barDirection != null && barDirection.equals("BAR"))
			bar.setBarDirection(BarDirection.BAR);
		else{
			bar.setBarDirection(BarDirection.COL);
		}
	}

	private void drawLineChartWithXDDF() {

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);
		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle(valueLabel[0]);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(values[0]);

		XDDFLineChartData chartData = (XDDFLineChartData) XDDFchart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
		chartData.setVaryColors(false);

		XDDFLineChartData.Series lineSeries = (XDDFLineChartData.Series) chartData.addSeries(cat, val);
		XDDFchart.plot(chartData);
		lineSeries.setMarkerStyle(MarkerStyle.NONE);
		lineSeries.setTitle(valueLabel[0], null);
		lineSeries.setSmooth(false);

		if(this.valueFormat != null){
			XDDFchart.getCTChart().getPlotArea().getValAxArray(0).addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getValAxArray(0).getNumFmt().setFormatCode(this.valueFormat);

			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);

			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().getNumFmt().setFormatCode(this.valueFormat);
		}
		XDDFchart.plot(chartData);
	}

	private void drawCombinedBarChartWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);

		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle(valueLabel[0] + " & " + valueLabel[1]);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val1 = XDDFDataSourcesFactory.fromArray(values[0]);
		XDDFNumericalDataSource<Double> val2 = XDDFDataSourcesFactory.fromArray(values[1]);

		XDDFChartData chartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		chartData.setVaryColors(true);
		XDDFChartData.Series series1 = chartData.addSeries(cat, val1);
		XDDFChartData.Series series2 = chartData.addSeries(cat, val2);

		series1.setTitle(valueLabel[0], null);
		series2.setTitle(valueLabel[1], null);

		CTChart ctChart = XDDFchart.getCTChart();
		CTBoolean ctboolean = CTBoolean.Factory.newInstance();

		ctboolean.setVal(true);
		ctChart.getPlotArea().getBarChartArray(0).addNewDLbls().setShowBubbleSize(ctboolean);
		ctboolean.setVal(false);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowVal(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowSerName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowPercent(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLegendKey(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowCatName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLeaderLines(ctboolean);

		XDDFchart.plot(chartData);

		XDDFBarChartData bar = (XDDFBarChartData) chartData;
		bar.setBarDirection(BarDirection.COL);
	}
	private void drawCombinedBarAndLineChartWithXDDF() {
		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val1 = XDDFDataSourcesFactory.fromArray(values[0]);
		XDDFNumericalDataSource<Double> val2 = XDDFDataSourcesFactory.fromArray(values[1]);
		XDDFNumericalDataSource<Double> val3 = XDDFDataSourcesFactory.fromArray(values[2]);

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);
		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle(valueLabel[0] + " & " + valueLabel[1]);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		XDDFChartData chartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		XDDFChartData.Series series1 = chartData.addSeries(cat, val1);
		XDDFChartData.Series series2 = chartData.addSeries(cat, val2);
		XDDFchart.plot(chartData);

		bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setVisible(false);
		XDDFValueAxis rightAxis = XDDFchart.createValueAxis(AxisPosition.RIGHT);
		rightAxis.setTitle(valueLabel[2]);
		rightAxis.setCrosses(AxisCrosses.MAX);
		bottomAxis.crossAxis(rightAxis);
		rightAxis.crossAxis(bottomAxis);
		XDDFChartData lineChartData = XDDFchart.createData(ChartTypes.LINE, bottomAxis, rightAxis);
		chartData.setVaryColors(true);
		XDDFLineChartData.Series lineSeries = (XDDFLineChartData.Series) lineChartData.addSeries(cat, val3);
		XDDFchart.plot(lineChartData);
		lineSeries.setMarkerStyle(MarkerStyle.NONE);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		//rightAxis.setCrosses(AxisCrosses.MAX);
		rightAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		series1.setTitle(valueLabel[0], null);
		series2.setTitle(valueLabel[1], null);
		lineSeries.setTitle(valueLabel[2], null);

		CTChart ctChart = XDDFchart.getCTChart();
		CTValAx valAx = ctChart.getPlotArea().getValAxArray(0); // get left axis
		valAx.addNewNumFmt().setFormatCode("#,##0.00\\ TL");


		if (ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr() == null)
			ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewSpPr();

		if (ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn() == null)
			ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().addNewLn();

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0)
				.getSpPr().getLn().setW(Units.pixelToEMU(3));

		if (ctChart.getPlotArea().getLineChartArray(0).getSerArray(0)
				.getSpPr().getLn().getSolidFill() == null)
			ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn().addNewSolidFill();

		if (ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn()
				.getSolidFill().getSrgbClr() == null)
			ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn()
					.getSolidFill().addNewSrgbClr();

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn()
				.addNewPrstDash().setVal(STPresetLineDashVal.SOLID);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewSmooth().setVal(false);

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0)
				.getSpPr().getLn().getSolidFill().getSrgbClr().setVal(new byte[]{(byte)0,(byte)0,(byte)255});

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);

		ctChart.getPlotArea().getValAxArray(0).getNumFmt().setSourceLinked(false);
		ctChart.getPlotArea().getValAxArray(0).getNumFmt().setFormatCode("#,##0.00\\ TL");

		CTBoolean ctboolean = CTBoolean.Factory.newInstance();

		ctboolean.setVal(true);
		ctChart.getPlotArea().getBarChartArray(0).addNewDLbls().setShowBubbleSize(ctboolean);
		ctboolean.setVal(false);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowVal(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowSerName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowPercent(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLegendKey(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowCatName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLeaderLines(ctboolean);
		XDDFBarChartData bar = (XDDFBarChartData) chartData;
		bar.setBarDirection(BarDirection.COL);
	}

	private void drawCombinedBarChartForDiscountWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);;
		//bottomAxis.setVisible(false);

		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle("TENZİLAT");
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);

		XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(values[0]);
		XDDFNumericalDataSource<Double> secondVal = XDDFDataSourcesFactory.fromArray(values[1]);
		XDDFNumericalDataSource<Double> thirdVal = XDDFDataSourcesFactory.fromArray(values[2]);

		XDDFChartData barChartDataPositive = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		barChartDataPositive.setVaryColors(true);
		XDDFChartData.Series positiveSeries = barChartDataPositive.addSeries(cat, val);

		XDDFChartData lineChartDataFirst = XDDFchart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
		XDDFLineChartData.Series firstLineSeries = (XDDFLineChartData.Series) lineChartDataFirst.addSeries(cat, secondVal);
		firstLineSeries.setMarkerStyle(MarkerStyle.NONE);

		XDDFChartData lineChartDataSecond = XDDFchart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
		XDDFLineChartData.Series secondLineSeries = (XDDFLineChartData.Series) lineChartDataSecond.addSeries(cat, thirdVal);
		secondLineSeries.setMarkerStyle(MarkerStyle.NONE);

		positiveSeries.setTitle(valueLabel[0], null);
		firstLineSeries.setTitle(valueLabel[1], null);
		secondLineSeries.setTitle(valueLabel[2], null);

		XDDFchart.plot(barChartDataPositive);
		XDDFchart.plot(lineChartDataFirst);
		XDDFchart.plot(lineChartDataSecond);

		CTChart ctChart = XDDFchart.getCTChart();
		CTValAx valAx = ctChart.getPlotArea().getValAxArray(0);
		valAx.addNewNumFmt().setFormatCode("0%");

		ctChart.getPlotArea().getBarChartArray(0).getSerArray(0).addNewInvertIfNegative().setVal(false);

		for(int i = 0; i < 2; i++){
			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).addNewSpPr();

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().addNewLn();

			ctChart.getPlotArea().getLineChartArray(i).getSerArray(0)
					.getSpPr().getLn().setW(Units.pixelToEMU(3));

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0)
					.getSpPr().getLn().getSolidFill() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn().addNewSolidFill();

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn()
					.getSolidFill().getSrgbClr() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn()
						.getSolidFill().addNewSrgbClr();
		}

		ctChart.getPlotArea().getLineChartArray(1).getSerArray(0).getSpPr().getLn().
				addNewPrstDash().setVal(STPresetLineDashVal.DASH);
		ctChart.getPlotArea().getLineChartArray(1).getSerArray(0).addNewSmooth().setVal(false);

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn()
				.addNewPrstDash().setVal(STPresetLineDashVal.SOLID);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewSmooth().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewNumFmt().setSourceLinked(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().getNumFmt().setFormatCode("%0.0");

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0)
				.getSpPr().getLn().getSolidFill().getSrgbClr().setVal(new byte[]{(byte)0,(byte)0,(byte)255});

		ctChart.getPlotArea().getLineChartArray(1).getSerArray(0)
				.getSpPr().getLn().getSolidFill().getSrgbClr().setVal(new byte[]{(byte)255,(byte)0,(byte)0});

		ctChart.getPlotArea().getValAxArray(0).getNumFmt().setSourceLinked(false);
		ctChart.getPlotArea().getValAxArray(0).getNumFmt().setFormatCode("%0.0");

		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);

		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewNumFmt().setSourceLinked(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().getNumFmt().setFormatCode("%0.0");

		XDDFBarChartData bar = (XDDFBarChartData) barChartDataPositive;
		bar.setBarDirection(BarDirection.COL);

	}

	private void drawCustomChartWithXDDF() {
		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);
		XDDFNumericalDataSource<Double> val = XDDFDataSourcesFactory.fromArray(values[0]);
		XDDFNumericalDataSource<Double> secondVal = XDDFDataSourcesFactory.fromArray(values[1]);

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle("YIL");
		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle("DOSYA SAYISI");
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		XDDFChartData lineChartData = XDDFchart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
		XDDFLineChartData.Series lineSeries = (XDDFLineChartData.Series) lineChartData.addSeries(cat, val);
		XDDFchart.plot(lineChartData);

		bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setVisible(false);
		XDDFValueAxis rightAxis = XDDFchart.createValueAxis(AxisPosition.RIGHT);
		rightAxis.setTitle("İHALE TUTARI (x1.000.000 TL)");
		rightAxis.setCrosses(AxisCrosses.MAX);
		bottomAxis.crossAxis(rightAxis);
		rightAxis.crossAxis(bottomAxis);
		XDDFChartData barChartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, rightAxis);
		barChartData.setVaryColors(true);
		XDDFChartData.Series series = barChartData.addSeries(cat, secondVal);
		XDDFchart.plot(barChartData);
		lineSeries.setMarkerStyle(MarkerStyle.NONE);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
		rightAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		series.setTitle("İHALE TUTARI (x1.000.000 TL)", null);
		lineSeries.setTitle("DOSYA SAYISI", null);

		CTChart ctChart = XDDFchart.getCTChart();
		CTValAx valAx = ctChart.getPlotArea().getValAxArray(1); // get right axis
		valAx.addNewNumFmt().setFormatCode("#,##0.00\\ TL");

		for(int i = 0; i < 1; i++){
			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).addNewSpPr();

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().addNewLn();

			ctChart.getPlotArea().getLineChartArray(i).getSerArray(0)
					.getSpPr().getLn().setW(Units.pixelToEMU(3));

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0)
					.getSpPr().getLn().getSolidFill() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn().addNewSolidFill();

			if (ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn()
					.getSolidFill().getSrgbClr() == null)
				ctChart.getPlotArea().getLineChartArray(i).getSerArray(0).getSpPr().getLn()
						.getSolidFill().addNewSrgbClr();
		}

		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getLineChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn().addNewPrstDash().setVal(STPresetLineDashVal.SOLID);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).addNewSmooth().setVal(false);
		ctChart.getPlotArea().getLineChartArray(0).getSerArray(0).getSpPr().getLn().getSolidFill().getSrgbClr().setVal(new byte[]{(byte)0,(byte)0,(byte)255});

		ctChart.getPlotArea().getValAxArray(1).getNumFmt().setSourceLinked(false); // right axis
		ctChart.getPlotArea().getValAxArray(1).getNumFmt().setFormatCode("#,##0.00\\ TL"); // right axis

		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).addNewDLbls().addNewShowVal().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);

		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewNumFmt().setSourceLinked(false);
		XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().getNumFmt().setFormatCode("#,##0.00\\ TL");

		XDDFBarChartData bar = (XDDFBarChartData) barChartData;
		bar.setBarDirection(BarDirection.COL);
	}

	private void drawStackedBarChartWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);

		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		//leftAxis.setTitle(valueLabel[0]);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);

		ArrayList<XDDFNumericalDataSource<Double>> vals = new ArrayList<>();
		for(int i = 0; i < values.length; i++){
			vals.add(XDDFDataSourcesFactory.fromArray(values[i]));
		}

		XDDFChartData chartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		chartData.setVaryColors(true);

		ArrayList<XDDFChartData.Series> series = new ArrayList<>();
		for(XDDFNumericalDataSource<Double> val : vals)
			series.add(chartData.addSeries(cat, val));

		for(int i = 0; i < valueLabel.length; i++)
			series.get(i).setTitle(valueLabel[i], null);


		XDDFchart.getCTChart().getPlotArea().getValAxArray(0).addNewNumFmt().setSourceLinked(false);
		XDDFchart.getCTChart().getPlotArea().getValAxArray(0).getNumFmt().setFormatCode("#,##0.00\\ TL");

		CTChart ctChart = XDDFchart.getCTChart();
		CTBoolean ctboolean = CTBoolean.Factory.newInstance();
		ctboolean.setVal(true);

		ctChart.getPlotArea().getBarChartArray(0).addNewDLbls().setShowVal(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLeaderLines(ctboolean);
		ctboolean.setVal(false);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowSerName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowPercent(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLegendKey(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowCatName(ctboolean);

		for(int i = 0; i < series.size(); i++){
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).addNewDLbls().addNewShowVal().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLeaderLines().setVal(true);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowSerName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowCatName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowPercent().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLegendKey().setVal(false);

			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().getNumFmt().setFormatCode("#,##0.00\\ TL");
		}

		XDDFchart.plot(chartData);

		XDDFBarChartData bar = (XDDFBarChartData) chartData;
		bar.setBarDirection(BarDirection.COL);
		bar.setBarGrouping(BarGrouping.STANDARD);

	}
	private void drawCombinedStackedBarChartWithXDDF() {

		XDDFChartLegend legend = XDDFchart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP);

		XDDFCategoryAxis bottomAxis = XDDFchart.createCategoryAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(categoryLabel);

		XDDFValueAxis leftAxis = XDDFchart.createValueAxis(AxisPosition.LEFT);
		//leftAxis.setTitle(valueLabel[0]);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

		XDDFDataSource<String> cat = XDDFDataSourcesFactory.fromArray(categories);

		ArrayList<XDDFNumericalDataSource<Double>> vals = new ArrayList<>();
		for(int i = 0; i < values.length; i++){
			vals.add(XDDFDataSourcesFactory.fromArray(values[i]));
		}

		XDDFChartData chartData = XDDFchart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
		chartData.setVaryColors(true);

		ArrayList<XDDFChartData.Series> series = new ArrayList<>();
		for(XDDFNumericalDataSource<Double> val : vals) {
			series.add(chartData.addSeries(cat, val));
		}
		for(int i = 0; i < valueLabel.length; i++)
			series.get(i).setTitle(valueLabel[i], null);

		XDDFchart.getCTChart().getPlotArea().getValAxArray(0).addNewNumFmt().setSourceLinked(false);
		XDDFchart.getCTChart().getPlotArea().getValAxArray(0).getNumFmt().setFormatCode("%0");

		CTChart ctChart = XDDFchart.getCTChart();
		ctChart.getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) 100);
		CTBoolean ctboolean = CTBoolean.Factory.newInstance();
		ctboolean.setVal(true);

		ctChart.getPlotArea().getBarChartArray(0).addNewDLbls().setShowVal(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowPercent(ctboolean);
		ctboolean.setVal(false);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLeaderLines(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowSerName(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowLegendKey(ctboolean);
		ctChart.getPlotArea().getBarChartArray(0).getDLbls().setShowCatName(ctboolean);

		for(int i = 0; i < series.size(); i++){
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).addNewDLbls().addNewShowVal().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLeaderLines().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowSerName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowCatName().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowPercent().setVal(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLegendKey().setVal(false);

			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewNumFmt().setSourceLinked(false);
			XDDFchart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().getNumFmt().setFormatCode("%0");
		}
		XDDFchart.plot(chartData);
		XDDFBarChartData bar = (XDDFBarChartData) chartData;
		bar.setBarDirection(BarDirection.COL);
		bar.setBarGrouping(BarGrouping.PERCENT_STACKED);

	}

	public JFreeChart getChart() { return this.chart; }

	public InputStream getInputStream() throws IOException {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		ChartUtils.writeChartAsJPEG(stream, this.chart, this.width, this.height);
		InputStream inputStream = new ByteArrayInputStream(stream.toByteArray());
		return inputStream;
	}
}
