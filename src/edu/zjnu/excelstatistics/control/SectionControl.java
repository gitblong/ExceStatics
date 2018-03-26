package edu.zjnu.excelstatistics.control;

import java.util.ArrayList;
import java.util.List;

import javax.swing.JTextField;

import org.apache.poi.ss.usermodel.Workbook;

import edu.zjnu.excelstatistics.bean.ValueBean;
import edu.zjnu.excelstatistics.service.SectionService;
import edu.zjnu.excelstatistics.utils.DrawChartLine;
import edu.zjnu.excelstatistics.utils.OutputExcelUtil;
import edu.zjnu.excelstatistics.utils.WDWUtil;

public class SectionControl {
	// excelFilePath, beginField, endField, sectionNum
	private String excelFilePath;
//	private JTextField beginField[];
//	private JTextField endField[];
	private int sectionNum;
	private ArrayList<Integer> mBeginList;
	private ArrayList<Integer> mEndList;
	public SectionControl() {

	}

//	public SectionControl(String excelFilePath, JTextField[] beginField, JTextField[] endField, int sectionNum) {
//		super();
//		this.excelFilePath = excelFilePath;
//		this.beginField = beginField;
//		this.endField = endField;
//		this.sectionNum = sectionNum;
//	}

	public SectionControl(String excelFilePath, ArrayList<Integer> mBeginList, ArrayList<Integer> mEndList,
			int sectionNum) {
		this.excelFilePath = excelFilePath;
		this.mBeginList = mBeginList;
		this.mEndList = mEndList;
		this.sectionNum = sectionNum;
	}

	public void getStatisticAllData(String excelPath) {
		WDWUtil wdwUtil = new WDWUtil();
		System.out.println(excelPath);
		System.out.println(wdwUtil.getWorkbook(excelPath) == null);
		SectionService sectionService = new SectionService(excelPath, mBeginList, mEndList, sectionNum);
		Workbook workbook = wdwUtil.getWorkbook(excelPath);
		List<List<ValueBean>> valueBeanLists = sectionService.getStatisticAllData(workbook);
		String headers[] = { "该段数据个数", "起始位置", "结束位置", "最小值", "最大值", "平均值", "数据条数" };

		OutputExcelUtil outputExcel = new OutputExcelUtil();
		outputExcel.fillAllExcelData(excelPath, valueBeanLists, workbook, headers);
		DrawChartLine scatterChart = new DrawChartLine(valueBeanLists, workbook, excelPath);
		scatterChart.drawAllTable(sectionNum);
	}

	public String getExcelFilePath() {
		return excelFilePath;
	}

	public void setExcelFilePath(String excelFilePath) {
		this.excelFilePath = excelFilePath;
	}


	public int getSectionNum() {
		return sectionNum;
	}

	public void setSectionNum(int sectionNum) {
		this.sectionNum = sectionNum;
	}

}
