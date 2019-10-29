package org.poi;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * POI读取excel有两种模式，一种是用户模式，一种是事件驱动模式 采用SAX事件驱动模式解决XLSX文件，可以有效解决用户模式内存溢出的问题， 该模式是POI官方推荐的读取大数据的模式， 在用户模式下，数据量较大，Sheet较多，或者是有很多无用的空行的情况下，容易出现内存溢出 用于解决.xlsx2007版本大数据量问题
 **/
public class ExcelXlsxReaderWithDefaultHandler extends DefaultHandler {

	/** 单元格中的数据可能的数据类型 */
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}

	/** 共享字符串表 */
	private SharedStringsTable sst;

	/** 上一次的索引值 */
	private StringBuilder lastIndex;

	/** 文件的绝对路径 */
	private String filePath = "";

	/** 工作表索引 */
	private int sheetIndex = 0;

	/** sheet名 */
	private String sheetName = "";

	/** 总行数 */
	private int totalRows = 0;

	/** 一行内cell集合 */
	private List<String> cellList = new ArrayList<String>();

	/** 判断整行是否为空行的标记 */
	private boolean flag = false;

	/** 当前行 */
	private int curRow = 1;

	/** 当前列 */
	private int curCol = 0;

	/** T元素标识 */
	private boolean isTElement;

	/** 判断上一单元格是否为文本空单元格 */
	private boolean charactersFlag = false;
	private boolean vtagFlag = false;

	/** 异常信息，如果为空则表示没有异常 */
	private String exceptionMessage;

	/** 单元格数据类型，默认为字符串类型 */
	private CellDataType nextDataType = CellDataType.SSTINDEX;

	private final DataFormatter formatter = new DataFormatter();

	/** 单元格日期格式的索引 */
	private short formatIndex;

	/** 日期格式字符串 */
	private String formatString;

	// 定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
	private String preRef = null, ref = null;

	// 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
	private String maxRef = null;

	/** 单元格 */
	private StylesTable stylesTable;

	/**
	 * 遍历工作簿中所有的电子表格 并缓存在mySheetList中
	 */
	public int process(String filename) throws Exception {
		filePath = filename;
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(pkg);
		stylesTable = xssfReader.getStylesTable();
		SharedStringsTable sst = xssfReader.getSharedStringsTable();
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		this.sst = sst;
		parser.setContentHandler(this);
		XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		while (sheets.hasNext()) { // 遍历sheet
			curRow = 1; // 标记初始行为第一行
			sheetIndex++;
			InputStream sheet = sheets.next(); // sheets.next()和sheets.getSheetName()不能换位置，否则sheetName报错
			sheetName = sheets.getSheetName();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource); // 解析excel的每条记录，在这个过程中startElement()、characters()、endElement()这三个函数会依次执行
			sheet.close();
		}
		return totalRows; // 返回该excel文件的总行数，不包括首列和空行
	}

	/** 第一个执行 */
	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// c => 单元格
		if ("c".equals(name)) {
			// 当前单元格的位置
			ref = attributes.getValue("r");

			// c标签开始时, 判断首行空值
			if (preRef == null) {
				int len = countNullCell(ref, "A") + 1;
				for (int i = 0; i < len; i++) {
					cellList.add(curCol, "");
					curCol++;
				}
				// 中间空值
			} else {
				int len = countNullCell(ref, preRef);
				for (int i = 0; i < len; i++) {
					cellList.add(curCol, "");
					curCol++;
				}
			}

			// 设定单元格类型
			this.setNextDataType(attributes);
			charactersFlag = false;
			vtagFlag = false;
		} else if ("v".equals(name)) {
			vtagFlag = true;
		}

		// 当元素为t时
		if ("t".equals(name)) {
			isTElement = true;
		} else {
			isTElement = false;
		}

		// 置空
		lastIndex = new StringBuilder();
	}

	/**
	 * 第二个执行 得到单元格对应的索引值或是内容值 如果单元格类型是字符串、INLINESTR、数字、日期，lastIndex则是索引值 如果单元格类型是布尔值、错误、公式，lastIndex则是内容值
	 */
	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		charactersFlag = true;
		lastIndex.append(new String(ch, start, length));
	}

	/**
	 * 第三个执行
	 */
	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		// t元素也包含字符串
		if (isTElement) {// 这个程序没经过
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			String value = lastIndex.toString();
			cellList.add(curCol, value);
			curCol++;
			isTElement = false;
			// 如果里面某个单元格含有值，则标识该行不为空行
			if (value != null && !"".equals(value)) {
				flag = true;
			}
		} else if ("c".equals(name)) {
			// c标签结束,判断是否存在v标签, 无v标签也要添加空值
			if (!vtagFlag) {
				cellList.add(curCol, "");
				curCol++;
			}
			vtagFlag = false;
			// c标签结束记录上一个列号
			preRef = ref;
		} else if ("v".equals(name)) {
			// v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
			// 根据索引值获取对应的单元格值
			String value = this.getDataValue(lastIndex.toString(), "");

			cellList.add(curCol, value);
			curCol++;
			// 如果里面某个单元格含有值，则标识该行不为空行
			if (value != null && !"".equals(value)) {
				flag = true;
			}
		} else {
			// 如果标签名称为row，这说明已到行尾，调用optRows()方法
			if ("row".equals(name)) {
				// 默认第一行为表头，以该行单元格数目为最大数目
				if (curRow == 1) {
					maxRef = ref;
				}
				// 补全一行尾部可能缺失的单元格
				if (maxRef != null) {
					int len = -1;
					// 前一单元格，true则不是文本空字符串，false则是文本空字符串
					if (charactersFlag) {
						len = countNullCell(maxRef, ref);
					} else {
						len = countNullCell(maxRef, preRef);
					}
					for (int i = 0; i <= len; i++) {
						cellList.add(curCol, "");
						curCol++;
					}
				}

				if (flag) { // 该行不为空行且该行不是第一行，则发送（第一行为列名，不需要）
					if (curRow == 1) {
						System.out.println(filePath + "\t" + sheetName + "\n" + cellList.size() + "\t" + Arrays.toString(cellList.toArray()));
					} else {
						ExcelReaderUtil.sendRows(filePath, sheetName, sheetIndex, curRow, cellList);
						totalRows++;
					}
				}

				cellList.clear();
				curRow++;
				curCol = 0;
				preRef = null;
				ref = null;
				flag = false;
			}
		}
	}

	/**
	 * 处理数据类型
	 */
	public void setNextDataType(Attributes attributes) {
		nextDataType = CellDataType.NUMBER; // cellType为空，则表示该单元格类型为数字
		formatIndex = -1;
		formatString = null;
		String cellType = attributes.getValue("t"); // 单元格类型
		String cellStyleStr = attributes.getValue("s"); //
		String columnData = attributes.getValue("r"); // 获取单元格的位置，如A1,B1

		if ("b".equals(cellType)) { // 处理布尔值
			nextDataType = CellDataType.BOOL;
		} else if ("e".equals(cellType)) { // 处理错误
			nextDataType = CellDataType.ERROR;
		} else if ("inlineStr".equals(cellType)) {
			nextDataType = CellDataType.INLINESTR;
		} else if ("s".equals(cellType)) { // 处理字符串
			nextDataType = CellDataType.SSTINDEX;
		} else if ("str".equals(cellType)) {
			nextDataType = CellDataType.FORMULA;
		}

		if (cellStyleStr != null) { // 处理日期
			int styleIndex = Integer.parseInt(cellStyleStr);
			XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
			formatIndex = style.getDataFormat();
			formatString = style.getDataFormatString();
			if (formatString.contains("m/d/yyyy") || formatString.contains("yyyy/mm/dd") || formatString.contains("yyyy/m/d")) {
				nextDataType = CellDataType.DATE;
				formatString = "yyyy-MM-dd hh:mm:ss";
			}

			if (formatString == null) {
				nextDataType = CellDataType.NULL;
				formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
			}
		}
	}

	/**
	 * 对解析出来的数据进行类型处理
	 * 
	 * @param value
	 *            单元格的值， value代表解析：BOOL的为0或1， ERROR的为内容值，FORMULA的为内容值，INLINESTR的为索引值需转换为内容值， SSTINDEX的为索引值需转换为内容值， NUMBER为内容值，DATE为内容值
	 * @param thisStr
	 *            一个空字符串
	 */
	public String getDataValue(String value, String thisStr) {
		switch (nextDataType) {
		// 这几个的顺序不能随便交换，交换了很可能会导致数据错误
		case BOOL: // 布尔值
			char first = value.charAt(0);
			thisStr = first == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR: // 错误
			thisStr = "\"ERROR:" + value.toString() + '"';
			break;
		case FORMULA: // 公式
			thisStr = '"' + value.toString() + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
			thisStr = rtsi.toString();
			rtsi = null;
			break;
		case SSTINDEX: // 字符串
			String sstIndex = value.toString();
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));// 根据idx索引值获取内容值
				thisStr = rtss.toString();
				// System.out.println(thisStr);
				// 有些字符串是文本格式的，但内容却是日期

				rtss = null;
			} catch (NumberFormatException ex) {
				thisStr = value.toString();
			}
			break;
		case NUMBER: // 数字
			if (formatString != null) {
				thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
			} else {
				thisStr = value;
			}
			thisStr = thisStr.replace("_", "").trim();
			break;
		case DATE: // 日期
			thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
			// 对日期字符串作特殊处理，去掉T
			thisStr = thisStr.replace("T", " ");
			break;
		default:
			thisStr = " ";
			break;
		}
		return thisStr;
	}

	public int countNullCell(String ref, String preRef) {
		// excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
		String xfd = ref.replaceAll("\\d+", "");
		String xfd_1 = preRef.replaceAll("\\d+", "");

		xfd = fillChar(xfd, 3, '@', true);
		xfd_1 = fillChar(xfd_1, 3, '@', true);

		char[] letter = xfd.toCharArray();
		char[] letter_1 = xfd_1.toCharArray();
		int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
		return res - 1;
	}

	public String fillChar(String str, int len, char let, boolean isPre) {
		int len_1 = str.length();
		if (len_1 < len) {
			if (isPre) {
				for (int i = 0; i < (len - len_1); i++) {
					str = let + str;
				}
			} else {
				for (int i = 0; i < (len - len_1); i++) {
					str = str + let;
				}
			}
		}
		return str;
	}

	public static void main(String[] args) throws Exception {
		new ExcelXlsxReaderWithDefaultHandler().process("C:/Users/Administrator/Desktop/test1.xlsx");
		// System.out.println(new ExcelXlsxReaderWithDefaultHandler().countNullCell("D", "B"));
	}
}
