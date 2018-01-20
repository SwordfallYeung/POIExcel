package org.poi;

import java.util.List;

/**
 * @author y
 * @create 2018-01-19 0:13
 * @desc
 **/
public class ExcelReaderUtil {
	//excel2003扩展名
	public static final String EXCEL03_EXTENSION = ".xls";
	//excel2007扩展名
	public static final String EXCEL07_EXTENSION = ".xlsx";

	/**
	 * 每获取一条记录，即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
	 *
	 * @param sheetName
	 * @param sheetIndex
	 * @param curRow
	 * @param cellList
	 */
	public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
			StringBuffer oneLineSb = new StringBuffer();
			oneLineSb.append(filePath);
			oneLineSb.append("--");
			oneLineSb.append("sheet" + sheetIndex);
			oneLineSb.append("::" + sheetName);//加上sheet名
			oneLineSb.append("--");
			oneLineSb.append("row" + curRow);
			oneLineSb.append("::");
			for (String cell : cellList) {
				oneLineSb.append(cell.trim());
				oneLineSb.append("|");
			}
			String oneLine = oneLineSb.toString();
			if (oneLine.endsWith("|")) {
				oneLine = oneLine.substring(0, oneLine.lastIndexOf("|"));
			}// 去除最后一个分隔符

			System.out.println(oneLine);
	}

	public static void readExcel(String fileName) throws Exception {
		int totalRows =0;
		if (fileName.endsWith(EXCEL03_EXTENSION)) { //处理excel2003文件
			ExcelXlsReader excelXls=new ExcelXlsReader();
			totalRows =excelXls.process(fileName);
		} else if (fileName.endsWith(EXCEL07_EXTENSION)) {//处理excel2007文件
			ExcelXlsxReader excelXlsxReader = new ExcelXlsxReader();
			totalRows = excelXlsxReader.process(fileName);
		} else {
			throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
		}
		System.out.println("发送的总行数：" + totalRows);
	}

	public static void main(String[] args) throws Exception {
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Date1\\H_20180111_Base_Date(4)_0420.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\REWORK\\H_20180105_Cto_REWORK_1600.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_ITEM\\H_20180116_ASS_ITEM_1915.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Process\\H_20180116_Cto_Process_2000.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Ship\\H_20180105_Cto_Ship_0005.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Po_sn\\H_20180105_Po_sn_1020.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_CODE\\H_20180109_Cto_CODE_1640.xlsx";*/
		String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_PRODUCT\\H_20171226_ASS_PRODUCT_0430 - 副本.xlsx";
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\MAC\\H_20180108_MAC_2130.xlsx";*/
		/*String path = "C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Data3\\H_20180106_Base_Data3_1520 - 副本.xlsx";*/
		/*String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\H_20171226_ASS_PRODUCT_0430.xls";*/
		ExcelReaderUtil.readExcel(path);
	}
}
