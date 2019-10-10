package org.poi;

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author y
 * @create 2018-01-19 14:18
 * @desc 用于解决.xls2003版本大数据量问题
 **/
public class ExcelXlsReader implements HSSFListener {

	private int minColums = -1;

	private POIFSFileSystem fs;

	/**
	 * 把第一行列名的长度作为列的总长
	 */
	private int totalColums = 0;

	/**
	 * 总行数
	 */
	private int totalRows=0;

	/**
	 * 上一行row的序号
	 */
	private int lastRowNumber;

	/**
	 * 上一单元格的序号
	 */
	private int lastColumnNumber;

	/**
	 * 是否输出formula，还是它对应的值
	 */
	private boolean outputFormulaValues = true;

	/**
	 * 用于转换formulas
	 */
	private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;

	//excel2003工作簿
	private HSSFWorkbook stubWorkbook;

	private SSTRecord sstRecord;

	private FormatTrackingHSSFListener formatListener;

	private final HSSFDataFormatter formatter = new HSSFDataFormatter();

	/**
	 * 文件的绝对路径
	 */
	private String filePath = "";

	//表索引
	private int sheetIndex = 0;

	private BoundSheetRecord[] orderedBSRs;

	@SuppressWarnings("unchecked")
	private ArrayList boundSheetRecords = new ArrayList();

	private int nextRow;

	private int nextColumn;

	private boolean outputNextStringRecord;

	//当前行
	private int curRow = 0;

	//存储一行记录所有单元格的容器
	private List<String> cellList = new ArrayList<String>();

	/**
	 * 判断整行是否为空行的标记
	 */
	private boolean flag = false;

	@SuppressWarnings("unused")
	private String sheetName;

	/**
	 * 遍历excel下所有的sheet
	 *
	 * @param fileName
	 * @throws Exception
	 */
	public int process(String fileName) throws Exception {
		filePath = fileName;
		this.fs = new POIFSFileSystem(new FileInputStream(fileName));
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
		formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, fs);

		return totalRows; //返回该excel文件的总行数，不包括首列和空行
	}

	/**
	 * HSSFListener 监听方法，处理Record
	 * 处理每个单元格
	 * @param record
	 */
	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String thisStr = null;
		String value = null;
		switch (record.getSid()) {
			case BoundSheetRecord.sid:
				boundSheetRecords.add(record);
				break;
			case BOFRecord.sid: //开始处理每个sheet
				BOFRecord br = (BOFRecord) record;
				if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
					//如果有需要，则建立子工作簿
					if (workbookBuildingListener != null && stubWorkbook == null) {
						stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
					}

					if (orderedBSRs == null) {
						orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
					}
					sheetName = orderedBSRs[sheetIndex].getSheetname();
					sheetIndex++;
				}
				break;
			case SSTRecord.sid:
				sstRecord = (SSTRecord) record;
				break;
			case BlankRecord.sid: //单元格为空白
				BlankRecord brec = (BlankRecord) record;
				thisRow = brec.getRow();
				thisColumn = brec.getColumn();
				thisStr = "";
				cellList.add(thisColumn, thisStr);
				break;
			case BoolErrRecord.sid: //单元格为布尔类型
				BoolErrRecord berec = (BoolErrRecord) record;
				thisRow = berec.getRow();
				thisColumn = berec.getColumn();
				thisStr = berec.getBooleanValue() + "";
				cellList.add(thisColumn, thisStr);
				checkRowIsNull(thisStr);  //如果里面某个单元格含有值，则标识该行不为空行
				break;
			case FormulaRecord.sid://单元格为公式类型
				FormulaRecord frec = (FormulaRecord) record;
				thisRow = frec.getRow();
				thisColumn = frec.getColumn();
				if (outputFormulaValues) {
					if (Double.isNaN(frec.getValue())) {
						outputNextStringRecord = true;
						nextRow = frec.getRow();
						nextColumn = frec.getColumn();
					} else {
						thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
					}
				} else {
					thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
				}
				cellList.add(thisColumn, thisStr);
				checkRowIsNull(thisStr);  //如果里面某个单元格含有值，则标识该行不为空行
				break;
			case StringRecord.sid: //单元格中公式的字符串
				if (outputNextStringRecord) {
					StringRecord srec = (StringRecord) record;
					thisStr = srec.getString();
					thisRow = nextRow;
					thisColumn = nextColumn;
					outputNextStringRecord = false;
				}
				break;
			case LabelRecord.sid:
				LabelRecord lrec = (LabelRecord) record;
				curRow = thisRow = lrec.getRow();
				thisColumn = lrec.getColumn();
				value = lrec.getValue().trim();
				value = value.equals("") ? "" : value;
				cellList.add(thisColumn, value);
				checkRowIsNull(value);  //如果里面某个单元格含有值，则标识该行不为空行
				break;
			case LabelSSTRecord.sid: //单元格为字符串类型
				LabelSSTRecord lsrec = (LabelSSTRecord) record;
				curRow = thisRow = lsrec.getRow();
				thisColumn = lsrec.getColumn();
				if (sstRecord == null) {
					cellList.add(thisColumn, "");
				} else {
					value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
					value = value.equals("") ? "" : value;
					cellList.add(thisColumn, value);
					checkRowIsNull(value);  //如果里面某个单元格含有值，则标识该行不为空行
				}
				break;
			case NumberRecord.sid: //单元格为数字类型
				NumberRecord numrec = (NumberRecord) record;
				curRow = thisRow = numrec.getRow();
				thisColumn = numrec.getColumn();

				//第一种方式
				//value = formatListener.formatNumberDateCell(numrec).trim();//这个被写死，采用的m/d/yy h:mm格式，不符合要求

				//第二种方式，参照formatNumberDateCell里面的实现方法编写
				Double valueDouble=((NumberRecord)numrec).getValue();
				String formatString=formatListener.getFormatString(numrec);
				if (formatString.contains("m/d/yy") || formatString.contains("yyyy/mm/dd") || formatString.contains("yyyy/m/d")){
					formatString="yyyy-MM-dd hh:mm:ss";
				}
				int formatIndex=formatListener.getFormatIndex(numrec);
				value=formatter.formatRawCellContents(valueDouble, formatIndex, formatString).trim();

				value = value.equals("") ? "" : value;
				//向容器加入列值
				cellList.add(thisColumn, value);
				checkRowIsNull(value);  //如果里面某个单元格含有值，则标识该行不为空行
				break;
			default:
				break;
		}

		//遇到新行的操作
		if (thisRow != -1 && thisRow != lastRowNumber) {
			lastColumnNumber = -1;
		}

		//空值的操作
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			curRow = thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			cellList.add(thisColumn, "");
		}

		//更新行和列的值
		if (thisRow > -1)
			lastRowNumber = thisRow;
		if (thisColumn > -1)
			lastColumnNumber = thisColumn;

		//行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (minColums > 0) {
				//列值重新置空
				if (lastColumnNumber == -1) {
					lastColumnNumber = 0;
				}
			}
			lastColumnNumber = -1;

			if (flag) { //该行不为空行且该行不是第一行，发送（第一行为列名，不需要）
				if (curRow == 0 ){
					totalColums = cellList.size(); //获取第一行列名的总数
				}else {
					//2003版尾部为空单元格的，xls里面是以该行最后一个有值的单元格为结束标记的，尾部空单元格跳过，故需补全
					if (cellList.size() <= totalColums){ // 其他行如果尾部单元格总数小于totalColums，则补全单元格
						for (int i = cellList.size(); i < totalColums; i++){
							cellList.add(i, "");
						}
					}
					ExcelReaderUtil.sendRows(filePath, sheetName, sheetIndex, curRow + 1, cellList); //每行结束时，调用sendRows()方法
					totalRows++;
				}
			}
			//清空容器
			cellList.clear();
			flag=false;
		}
	}

	/**
	 * 如果里面某个单元格含有值，则标识该行不为空行
	 * @param value
	 */
	public void checkRowIsNull(String value){
		if (value != null && !"".equals(value)) {
			flag = true;
		}
	}
}