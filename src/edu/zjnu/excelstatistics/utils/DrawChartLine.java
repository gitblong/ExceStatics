package edu.zjnu.excelstatistics.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;

import edu.zjnu.excelstatistics.bean.ValueBean;

/**
 * Illustrates how to create a simple scatter chart.
 * 
 */
public class DrawChartLine {
	private String mFilePath;
	final int COLNUM = 14;
	private Sheet sheet;
	private Workbook mWorkbook;
	private List<List<ValueBean>> mValueBeanLists;

	public DrawChartLine() {
	}

	public DrawChartLine(List<List<ValueBean>> mValueBeanLists, Workbook mWorkbook, String mFilePath) {
		this.mValueBeanLists = mValueBeanLists;
		this.mWorkbook = mWorkbook;
		this.mFilePath = mFilePath;
	}

	public void init(int sheetId) {
		WDWUtil wdwUtil = new WDWUtil();
		mWorkbook = wdwUtil.getWorkbook(mFilePath);
		sheet = mWorkbook.getSheetAt(sheetId);
	}

	public void createIndex(int dataSize) {
		int value = 0;
		/*
		 * 创建数据下标的 行 起始位置 rowIndex
		 * 
		 */
		for (int rowIndex = 1; rowIndex < dataSize; rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			Cell cell = row.createCell(1);
			cell.setCellValue(value++);
		}
	}

	public int createLineXY(int array[], int rowNum) {
		int yValue[] = { 0, 5 };
		for (int j = 0; j < array.length; j++) {
			for (int index = 0; index <= 1; index++) {
				Row row = sheet.getRow(rowNum++);
				int cellIndex = COLNUM;
				for (int i = 0; i < yValue.length; i++) {
					Cell cell = row.createCell(cellIndex++);
					switch (i) {
					case 0:
						cell.setCellValue(array[j]);
						break;
					case 1:
						cell.setCellValue(yValue[index]);
						break;
					}
				}
			}
		}
		return rowNum;
	}

	public void drawAllTable(int sectionNums) {
		int dataSize = mValueBeanLists.get(0).get(0).getTotal();
		int begin[] = new int[sectionNums];
		int end[] = new int[sectionNums];
		int sheetSizes = mValueBeanLists.size();
		int sheetNum = 1;
		for (List<ValueBean> list : mValueBeanLists) {
			int sectionNum = 1;
			System.out.println("----------------第" + sheetNum++
					+ "个Sheet页------------------------------------------------------------------------");
			int beginflag = 0;
			int endflag = 0;
			for (ValueBean valueBean : list) {
				begin[beginflag++] = valueBean.getBeginIndex();
				end[endflag++] = valueBean.getEndIndex();
				System.out.println("第" + (sectionNum++) + "段" + valueBean.toString());
			}
		}
		for (int sheetId = 0; sheetId < sheetSizes; sheetId++) {
			try {
				drawTable(dataSize, begin, end, sheetId);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	public void drawTable(int dataSize, int begin[], int end[], int sheetId) throws IOException {
		init(sheetId);

		/*
		 * 
		 * 创建横线数据 行 的起始位置
		 * 
		 * @rowNum
		 * 
		 */
		int rowNum = 43;
		rowNum = createLineXY(begin, rowNum);
		createLineXY(end, rowNum);
		createIndex(dataSize);

		Drawing drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 14, 27);

		Chart chart = drawing.createChart(anchor);
		ChartLegend legend = chart.getOrCreateLegend();
		legend.setPosition(LegendPosition.TOP_RIGHT);
		// legend.setPosition(LegendPosition.RIGHT);
		ScatterChartData data = chart.getChartDataFactory().createScatterChartData();
		ValueAxis bottomAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
		ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

		ChartDataSource<Number> dataX = DataSources.fromNumericCellRange(sheet,
				new org.apache.poi.ss.util.CellRangeAddress(0, 1339, 1, 1));
		ChartDataSource<Number> dataY = DataSources.fromNumericCellRange(sheet,
				new org.apache.poi.ss.util.CellRangeAddress(0, 1339, 0, 0));
		data.addSerie(dataX, dataY);
		int row = 43;
		for (int i = 0; i < begin.length; i++) {
			ChartDataSource<Number> beginX = DataSources.fromNumericCellRange(sheet,
					new org.apache.poi.ss.util.CellRangeAddress(row, row + 1, COLNUM, COLNUM));
			ChartDataSource<Number> beginY = DataSources.fromNumericCellRange(sheet,
					new org.apache.poi.ss.util.CellRangeAddress(row, row + 1, COLNUM + 1, COLNUM + 1));
			row += 2;
			data.addSerie(beginX, beginY);

		}
		for (int i = 0; i < end.length; i++) {
			ChartDataSource<Number> endX = DataSources.fromNumericCellRange(sheet,
					new org.apache.poi.ss.util.CellRangeAddress(row, row + 1, COLNUM, COLNUM));
			ChartDataSource<Number> endY = DataSources.fromNumericCellRange(sheet,
					new org.apache.poi.ss.util.CellRangeAddress(row, row + 1, COLNUM + 1, COLNUM + 1));
			row += 2;
			data.addSerie(endX, endY);
		}

		chart.plot(data, bottomAxis, leftAxis);

		// Write the output to a file
		File file = new File(mFilePath);
		FileOutputStream fileOut = new FileOutputStream(file);
		mWorkbook.write(fileOut);
		fileOut.close();
	}

	public Workbook getmWorkbook() {
		return mWorkbook;
	}

	public void setmWorkbook(Workbook mWorkbook) {
		this.mWorkbook = mWorkbook;
	}
}