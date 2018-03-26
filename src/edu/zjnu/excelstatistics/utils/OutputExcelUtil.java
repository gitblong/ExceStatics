package edu.zjnu.excelstatistics.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import edu.zjnu.excelstatistics.bean.ValueBean;

public class OutputExcelUtil {
	/** 错误信息 */
	private String errorInfo;
	private Workbook mWorkbook;

	public Workbook getmWorkbook() {
		return mWorkbook;
	}

	public void setmWorkbook(Workbook mWorkbook) {
		this.mWorkbook = mWorkbook;
	}

	/**
	 * @描述：验证excel文件
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：boolean
	 */

	public boolean validateExcel(String filePath) {
		/** 检查文件名是否为空或者是否是Excel格式的文件 */
		if (filePath == null || !(WDWUtil.isExcel2003(filePath) || WDWUtil.isExcel2007(filePath))) {
			errorInfo = "文件名不是excel格式";
			return false;
		}
		/** 检查文件是否存在 */
		File file = new File(filePath);
		if (file == null || !file.exists()) {
			errorInfo = "文件不存在";
			return false;
		}
		return true;
	}

	/**
	 * @描述：根据文件名读取excel文件
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：List
	 */
	public void filleExcel(String filePath, List<List<ValueBean>> valueBeanLists, String[] headers) {
		InputStream is = null;
		try {
			/** 验证文件是否合法 */
			if (!validateExcel(filePath)) {
				System.out.println(errorInfo);
				return;
			}
			/** 判断文件的类型，是2003还是2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(filePath)) {
				isExcel2003 = false;
			}
			/** 调用本类提供的根据流读取的方法 */
			File file = new File(filePath);
			is = new FileInputStream(file);
			fillExcel(filePath, is, isExcel2003, valueBeanLists, headers);
			is.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					is = null;
					e.printStackTrace();
				}
			}
		}
		/** 返回最后读取的结果 */
		return;
	}

	/**
	 * @描述：根据流读取Excel文件
	 * @参数：@param inputStream
	 * @参数：@param isExcel2003
	 * @参数：@return
	 * @返回值：List
	 */
	public void fillExcel(String filePath, InputStream inputStream, boolean isExcel2003,
			List<List<ValueBean>> valueBeanLists, String[] headers) {
		mWorkbook = null;
		try {
			/** 根据版本选择创建Workbook的方式 */
			if (isExcel2003) {
				mWorkbook = new HSSFWorkbook(inputStream);
			} else {
				mWorkbook = new XSSFWorkbook(inputStream);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
		fillAllExcelData(filePath, valueBeanLists, mWorkbook, headers);
	}

	public void fillAllExcelData(String filePath, List<List<ValueBean>> valueBeanLists, Workbook wb, String[] headers) {
		int flag = 0;
		for (List<ValueBean> valueBeanList : valueBeanLists) {
			fillExcelData(filePath, valueBeanList, wb, headers, flag++);
		}
	}

	public void fillExcelData(String filePath, List<ValueBean> valueBeanList, Workbook workbook, String[] headers,
			int sheetFlag) {
		edu.zjnu.excelstatistics.utils.WDWUtil wdwUtil = new edu.zjnu.excelstatistics.utils.WDWUtil();
		workbook = wdwUtil.getWorkbook(filePath);
		/*
		 * 
		 * 改变数据输出      行      的起始位置
		 * 
		 * rowIndex
		 * 
		 */
		int rowIndex = 31;
		int cellFlag = 1;
		Sheet sheet = workbook.getSheetAt(sheetFlag);
		Row row = sheet.getRow(rowIndex++);
		for (int i = 0; i < headers.length; i++) {
			/*
			 * 
			 * 改变显示第几段数据的    表头    的    列    的起始位置（i + 3）
			 * (i+3)
			 * 
			 */
			row.createCell(i + 3).setCellValue(headers[i]);
		}
		for (ValueBean valueBean : valueBeanList) {
			/*
			 * 
			 * 改变   数据值     列      的起始位置
			 * cellIndex
			 * 
			 * 
			 */
			int cellIndex = 2;
			row = sheet.getRow(rowIndex++);
			row.createCell(cellIndex++).setCellValue("第" + (cellFlag++) + "段");
			row.createCell(cellIndex++).setCellValue(valueBean.getCount());
			row.createCell(cellIndex++).setCellValue(valueBean.getBeginIndex());
			row.createCell(cellIndex++).setCellValue(valueBean.getEndIndex());
			row.createCell(cellIndex++).setCellValue(valueBean.getMax());
			row.createCell(cellIndex++).setCellValue(valueBean.getMin());
			row.createCell(cellIndex++).setCellValue(valueBean.getAverage());
			row.createCell(cellIndex++).setCellValue(valueBean.getTotal());
		}

		try {
			FileOutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
