package com.paraboly.reportlib;

import lombok.SneakyThrows;
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

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
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

	private DefaultCategoryDataset dataset;
	private String title, categoryLabel, valueLabel;
	private JFreeChart chart;

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

	public ChartDrawingService(String title, String categoryLabel, String valueLabel){
		dataset = new DefaultCategoryDataset();
		this.title = title;
		this.categoryLabel = categoryLabel;
		this.valueLabel = valueLabel;
	}

	public ChartDrawingService addData(List<?> dataList, String categoryMethod, String valueMethod, String groupName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
		for (Object data: dataList) {
			String category = (String) data.getClass().getMethod(categoryMethod).invoke(data);
			Float value = Float.valueOf(data.getClass().getMethod(valueMethod).invoke(data).toString());
			dataset.addValue(value, groupName, category);
		}
		return this;
	}

	@SneakyThrows
	public void draw(String chartType) {
		switch (chartType) {
			case "bar":
				chart = drawBarChart();
				break;
			case "pie":
				chart = drawPieChart();
				break;
			default:
				break;
		}
	}

	private JFreeChart drawBarChart() {
		JFreeChart barChartObject = ChartFactory.createBarChart(
				title,categoryLabel,
				valueLabel,
				this.dataset, PlotOrientation.VERTICAL,
				true,true,false);
		barChartObject.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions.UP_45);
		barChartObject.getCategoryPlot().getDomainAxis().setMaximumCategoryLabelLines(2);
		barChartObject.getPlot().setBackgroundPaint(Color.white);
		barChartObject.getPlot().setDrawingSupplier(new ChartDrawingSupplier());
		return barChartObject;
	}

	private JFreeChart drawPieChart() {
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

	public JFreeChart getChart() { return this.chart; }

	public InputStream getInputStream() throws IOException {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		ChartUtils.writeChartAsJPEG(stream, this.chart, this.width, this.height);
		InputStream inputStream = new ByteArrayInputStream(stream.toByteArray());
		return inputStream;
	}
}
