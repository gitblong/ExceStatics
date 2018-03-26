package edu.zjnu.excelstatistics.service;

import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.JTextField;

import org.apache.poi.ss.usermodel.Workbook;

import edu.zjnu.excelstatistics.ExcelDao.ImportExeclDao;
import edu.zjnu.excelstatistics.bean.ValueBean;

public class SectionService {
	
	private String excelPath;
//	private JTextField[] beginTextField;
//	private JTextField[] endTextField;
	private int sectionNum;
	private ArrayList<Integer> mBeginList;
	private ArrayList<Integer> mEndList;

//	public SectionService(String excelPath, JTextField[] beginTextField, JTextField[] endTextField, int sectionNum) {
//		this.excelPath = excelPath;
//		this.beginTextField = beginTextField;
//		this.endTextField = endTextField;
//		this.sectionNum = sectionNum;
//
//	}
	public SectionService(String excelPath, ArrayList<Integer> mBeginList, ArrayList<Integer> mEndList,
			int sectionNum) {
		this.excelPath = excelPath;
		this.mBeginList = mBeginList;
		this.mEndList = mEndList;
		this.sectionNum = sectionNum;
	}
	public List<List<ValueBean>> getStatisticAllData(Workbook workbook){
		ImportExeclDao importExecl2 = new ImportExeclDao();
		List<List<List<String>>> list = importExecl2.read(workbook);
		System.out.println("getStatisticAllData-list.size-"+list.size());
		List<List<ValueBean>> valueLists = new ArrayList<List<ValueBean>>();
		if (list!= null) {
			
			for (int i = 0; i < list.size(); i++) {
				List<ValueBean> valueList = new ArrayList<ValueBean>();
				List<List<String>> sheetList = list.get(i);
				List<Double> cellDoubleList = new ArrayList<Double>();
				for (int j = 1; j < sheetList.size(); j++) {
					List<String> cellList = sheetList.get(j);
					for (int k = 0; k < cellList.size(); k++) {
						String cellString = cellList.get(k);
						double cellDouble = Double.parseDouble(cellString);
						cellDoubleList.add(cellDouble);
					}
				}
				valueList = getStatisticData(cellDoubleList,i+1);
				valueLists.add(valueList);
			}
		}
		return valueLists;
	}

	public List<ValueBean> getStatisticData(List<Double> dataList,int n) {
		List<ValueBean> valueBeanList = new ArrayList<ValueBean>();
		if (dataList.size()!=0) {
			
			for (int i = 0; i < sectionNum; i++) {
				int begin = 50 * mBeginList.get(i);
				int end = 50 * mEndList.get(i);
				int subLength = end - begin;
				if (begin <= dataList.size()) {
					if (end > dataList.size()) {
						subLength = dataList.size() - begin;
						end = dataList.size();
//						System.out.println("---------if (begin <= dataArray.length) {begin"+begin+"dataList"+dataList.size());
					}
				} else {
//					System.out.println("---------else (begin > dataArray.length) {"+begin+"dataList"+dataList.size());
					JOptionPane.showMessageDialog(null, "第"+n+"个sheet表中的第" + (i+1) + "段已无数据");
					break;
				}
				double sectionData[] = new double[subLength];
				int k = 0;
//				System.out.println("begin"+begin+"end"+end+"dataList"+dataList.size());
				for (int j = begin; j < end; j++) {
//					System.out.println("dataList-------getStatisticData------"+dataList.size());
					sectionData[k] = dataList.get(j);
					k++;
				}
				double maxValue = getMax(sectionData);
				double minValue = getMin(sectionData);
				double averageValue = getAverage(sectionData);
				ValueBean valueBean = new ValueBean(subLength,begin,end,minValue, maxValue, averageValue,dataList.size());
				valueBeanList.add(valueBean);
			}
		}
		return valueBeanList;
	}

	private double getAverage(double[] sectionData) {
		double average = 0;
		double sum = 0;
		for (int i = 0; i < sectionData.length; i++) {
			sum += sectionData[i];
		}
		average = sum / (sectionData.length);
		return (int)(average*1000)/1000.0;
	}

	private double getMax(double[] sectionData) {
		double max = sectionData[0];
		for (int i = 1; i < sectionData.length; i++) {
			if (sectionData[i] > max) {
				max = sectionData[i];
			}
		}
		return (int)(max*1000)/1000.0;
	}

	private double getMin(double[] sectionData) {
		double min = sectionData[0];
		for (int i = 1; i < sectionData.length; i++) {
			if (sectionData[i] < min) {
				min = sectionData[i];
			}
		}
		return (int)(min*1000)/1000.0;
	}


}
