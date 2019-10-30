package excel;

import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Arrays;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import com.JdbcUtil;

/**
 * 使用CVS模式解决XLSX文件，可以有效解决用户模式内存溢出的问题
 */
public class ExcelXlsxSAX extends DefaultHandler {
	/** 类型枚举 */
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
	}

	private static String sheetName; // Sheet页名称

	private final static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 日期输出格式
	private final DataFormatter formatter = new DataFormatter();// 日期格式化取值

	private StylesTable styles;// 格式化样式表
	private ReadOnlySharedStringsTable sst;// 共享字符串表

	private int totalRows = 0;// 单页总行数
	private int colNo = -1;// 当前列

	private short formatIndex;// 单元格日期格式的索引
	private String formatString;// 日期格式字符串

	private CellDataType nextDataType;// 单元格数据类型
	private StringBuffer valueTemp = new StringBuffer();// 单元格数值缓存

	private String[] colls;// 每行数据

	/** 初始化处理程序 */
	public void process(String path, int minColumnCount) {
		colls = new String[minColumnCount];
		OPCPackage pkg = null;
		try {
			// 1.根据Excel获取OPCPackage对象
			pkg = OPCPackage.open(path, PackageAccess.READ);
			// 2.创建XSSFReader对象
			XSSFReader xssfReader = new XSSFReader(pkg);
			// 3.获取SharedStringsTable对象
			sst = new ReadOnlySharedStringsTable(pkg);
			// 4.获取StylesTable对象
			styles = xssfReader.getStylesTable();
			// 5.创建获取解析器Sax的XmlReader对象
			// XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
			XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
			// 6.设置处理器
			parser.setContentHandler(this);
			// 7.遍历Sheet页
			XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
			while (sheets.hasNext()) {
				totalRows = 0;
				InputStream stream = sheets.next();
				sheetName = sheets.getSheetName();
				try {
					// 解析excel，依次调用startElement()、characters()、endElement()
					parser.parse(new InputSource(stream));
				} finally {
					stream.close();
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (OpenXML4JException e) {
			e.printStackTrace();
		} catch (SAXException e) {
			e.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} finally {
			if (pkg != null) {
				try {
					pkg.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// 数值标签, 初始化缓存的值
		if ("inlineStr".equals(name) || "v".equals(name)) {
			valueTemp.setLength(0);
			// 单元格标签
		} else if ("c".equals(name)) {
			// 取列标识
			String r = attributes.getValue("r");
			int firstDigit = -1;
			for (int c = 0; c < r.length(); ++c) {
				// 判断字符是否为数字
				if (Character.isDigit(r.charAt(c))) {
					firstDigit = c;
					break;
				}
			}
			// 取列角标
			colNo = nameToColumn(r.substring(0, firstDigit));
			// 重置格式化值
			this.formatIndex = -1;
			this.formatString = null;
			this.nextDataType = CellDataType.NUMBER;
			// 数值类型
			String cellType = attributes.getValue("t");
			// 单元格格式
			String cellStyleStr = attributes.getValue("s");
			if ("b".equals(cellType)) {
				nextDataType = CellDataType.BOOL;
			} else if ("e".equals(cellType)) {
				nextDataType = CellDataType.ERROR;
			} else if ("inlineStr".equals(cellType)) {
				nextDataType = CellDataType.INLINESTR;
			} else if ("s".equals(cellType)) {
				nextDataType = CellDataType.SSTINDEX;
			} else if ("str".equals(cellType)) {
				nextDataType = CellDataType.FORMULA;
			} else if (cellStyleStr != null) {
				XSSFCellStyle style = styles.getStyleAt(Integer.parseInt(cellStyleStr));
				this.formatIndex = style.getDataFormat();
				this.formatString = style.getDataFormatString();
				// 日期格式
				if (formatString.contains("yy")) {
					formatString = "yyyy-MM-dd hh:mm:ss";
				}
				if (this.formatString == null) {
					this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
				}
			}
		}
	}

	/** 获取列角标 */
	private int nameToColumn(String name) {
		int column = -1;
		for (int i = 0; i < name.length(); ++i) {
			int c = name.charAt(i);
			column = (column + 1) * 26 + c - 'A';
		}
		return column;
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		valueTemp.append(ch, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		// 数值标签
		if ("v".equals(name)) {
			// 将数据添加到行数组
			colls[colNo] = getValue(valueTemp, nextDataType);
			// 遇见行结束标签, 则写出一行数据
		} else if ("row".equals(name)) {
			// 输出首行
			if (totalRows == 0) {
				System.out.println(Arrays.toString(colls));
			} else {
				sendRows(totalRows, colls);
			}
			totalRows++;
			// 清空行数组
			for (int i = 0; i < colls.length; i++) {
				colls[i] = null;
			}
		}
	}

	/** 根据单元格类型取值 */
	private String getValue(StringBuffer valueTemp, CellDataType nextDataType) {
		String value = "";
		switch (nextDataType) {
		case BOOL:
			value = valueTemp.charAt(0) == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR:
			break;
		case FORMULA:
			// 获取公式类型值
			value = valueTemp.toString();
			break;
		case INLINESTR:
			value = new XSSFRichTextString(valueTemp.toString()).toString();
			break;
		case SSTINDEX:
			// 获取文本类型值
			try {
				value = new XSSFRichTextString(sst.getEntryAt(Integer.parseInt(valueTemp.toString()))).toString();
			} catch (NumberFormatException ex) {
				ex.printStackTrace();
			}
			break;
		case NUMBER:
			// 判断是否是日期格式
			if (HSSFDateUtil.isADateFormat(this.formatIndex, valueTemp.toString())) {
				value = sdf.format(HSSFDateUtil.getJavaDate(Double.parseDouble(valueTemp.toString())));
			} else if (this.formatString != null) {
				value = formatter.formatRawCellContents(Double.parseDouble(valueTemp.toString()), this.formatIndex, this.formatString);
			} else {
				value = valueTemp.toString();
			}
			break;
		default:
			value = valueTemp.toString();
			break;
		}
		return "".equals(value.trim()) ? null : value.trim();
	}

	public static void sendRows(int curRow, String[] cells) {
		System.out.println(cells.length + "\t" + Arrays.toString(cells));
		/*System.out.println(sheetName + "\t" + "第 " + curRow + " 行");
		try {
			jdbcUtil.update(sql, cells);
		} catch (SQLException e) {
			e.printStackTrace();
		}*/
	}

	public static void main(String[] args) throws Exception {
		try {
			new ExcelXlsxSAX().process("C:/Users/Administrator/Desktop/test.xlsx", 4);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
