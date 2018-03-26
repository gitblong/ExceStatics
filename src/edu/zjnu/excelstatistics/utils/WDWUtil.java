package edu.zjnu.excelstatistics.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WDWUtil {

	String errorInfo;

	public String getErrorInfo() {
		return errorInfo;
	}

	public void setErrorInfo(String errorInfo) {
		this.errorInfo = errorInfo;
	}

	/**
	 * @描述：是否是2003的excel，返回true是2003
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：boolean
	 */

	public static boolean isExcel2003(String filePath) {
		return filePath.matches("^.+\\.(?i)(xls)$");
	}

	/**
	 * @描述：是否是2007的excel，返回true是2007
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：boolean
	 */
	public static boolean isExcel2007(String filePath) {
		return filePath.matches("^.+\\.(?i)(xlsx)$");
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
	 * @描述：根据文件名读取workbook
	 * @参数：@param filePath 文件完整路径
	 * @参数：@return
	 * @返回值：List
	 */
	public Workbook getWorkbook(String filePath) {
		Workbook workbook = null;
		InputStream inputStream = null;
		try {
			/** 验证文件是否合法 */
			if (!validateExcel(filePath)) {
				System.out.println(errorInfo);
				return null;
			}
			/** 判断文件的类型，是2003还是2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(filePath)) {
				isExcel2003 = false;
			}
			/** 调用本类提供的根据流读取的方法 */
			File file = new File(filePath);
			inputStream = new FileInputStream(file);
			/** 根据版本选择创建Workbook的方式 */
			// POIFSFileSystem fs = new POIFSFileSystem(inputStream);
			workbook = new XSSFWorkbook(inputStream);
			
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					inputStream = null;
					e.printStackTrace();
				}
			}
		}
		/** 返回最后读取的结果 */
		return workbook;
	}
}
