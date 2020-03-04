package com.paraboly.reportlib;

import lombok.SneakyThrows;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.util.TableOrder;
import org.jfree.data.category.CategoryToPieDataset;
import org.jfree.data.category.DefaultCategoryDataset;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

public class ChartDrawingService {
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
				throw new Exception("Chart type unsupported");
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
		return barChartObject;
	}

	private JFreeChart drawPieChart() {
		CategoryToPieDataset pieDataset = new CategoryToPieDataset(dataset, TableOrder.BY_ROW, 0);
		JFreeChart pieChartObject = ChartFactory.createPieChart(title, pieDataset);
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
