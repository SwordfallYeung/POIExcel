package org.poi;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

/**
 * @author y15079
 * @create 2018-01-24 13:56
 * @desc
 * POI读取excel有两种模式，一种是用户模式，一种是事件驱动模式
 * 采用SAX事件驱动模式解决XLSX文件，可以有效解决用户模式内存溢出的问题，
 * 该模式是POI官方推荐的读取大数据的模式，
 * 在用户模式下，数据量较大，Sheet较多，或者是有很多无用的空行的情况下，容易出现内存溢出
 * <p>
 * 用于解决.xlsx2007版本大数据量问题
 * 这里调用的是SheetContentsHandler的API
 *
 * 估计是封装DefaultHandler的，效率比不上DefaultHandler
 **/
public class ExcelXlsxReaderWithSheetContentsHandler {
	private String filename;
	private XSSFSheetXMLHandler.SheetContentsHandler handler;

	public ExcelXlsxReaderWithSheetContentsHandler(String filename) {
		this.filename = filename;
	}

	public ExcelXlsxReaderWithSheetContentsHandler setHandler(XSSFSheetXMLHandler.SheetContentsHandler handler){
		this.handler=handler;
		return this;
	}

	public void parse(){
		OPCPackage pkg=null;
		InputStream sheetInputStream=null;
		try {
			pkg=OPCPackage.open(filename, PackageAccess.READ);
			XSSFReader xssfReader=new XSSFReader(pkg);
			StylesTable styles=xssfReader.getStylesTable();
			ReadOnlySharedStringsTable strings=new ReadOnlySharedStringsTable(pkg);
			sheetInputStream=xssfReader.getSheetsData().next();
			processSheet(styles, strings, sheetInputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			if (sheetInputStream!=null){
				try {
					sheetInputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (pkg!=null){
				try {
					pkg.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	private void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings,InputStream sheetInputStream ) throws Exception{
		XMLReader sheetParser= SAXHelper.newXMLReader();
		if (handler!=null){
			sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles,strings,handler,false));
		}else {
			sheetParser.setContentHandler(new XSSFSheetXMLHandler(styles,strings,new SimpleSheetContentsHandler(),false));
		}
		sheetParser.parse(new InputSource(sheetInputStream));
	}

	public static class SimpleSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler{
		protected List<String> row =new LinkedList<String>();

		public void startRow(int rowNum) {
			row .clear();
		}

		public void endRow(int rowNum) {
			System.err.println(rowNum + " : " + row );
		}

		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			row .add(formattedValue);
		}

		public void headerFooter(String text, boolean isHeader, String tagName) {

		}
	}

	public static void readXlsx(String filePath){
		long start=System.currentTimeMillis();
		final List<List<String>> table = new ArrayList<List<String>>();
		new ExcelXlsxReaderWithSheetContentsHandler(filePath).setHandler(new ExcelXlsxReaderWithSheetContentsHandler.SimpleSheetContentsHandler(){
			private List<String> fields;

			@Override
			public void endRow(int rowNum) {
				if (rowNum == 0){
					// 第一行中文描述忽略
				}else {
					//数据
					table.add(row);
					System.out.println(row);
				}
			}
		}).parse();

		for (List<String> tab : table){

		}

		long end=System.currentTimeMillis();

		System.err.println(table.size());
		System.err.println(end - start);
	}
}
