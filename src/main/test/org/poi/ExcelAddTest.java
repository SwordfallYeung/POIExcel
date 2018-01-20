package org.poi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author y15079
 * @create 2018-01-15 16:37
 * @desc
 **/
public class ExcelAddTest {
	public static void main(String[] args) throws Exception{
		createAndCopyFile();

	}

	public static void createAndCopyFile()throws Exception{
		Map<String,String> fileMap=new HashMap<String, String>();

		String path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_ITEM\\H_20180116_ASS_ITEM_1915.xlsx";
		String targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_ITEM\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_ASS_ITEM_"+timeFormat()+".xlsx";
//		fileMap.put(path,targetPath);

		path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Date1\\H_20180111_Base_Date(4)_0420.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Date1\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Base_Date(4)_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);

		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Po_sn\\H_20180105_Po_sn_1020.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Po_sn\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Po_sn_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

//		path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Ship\\H_20180105_Cto_Ship_0005.xlsx";
//		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Ship\\";
//		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Cto_Ship_"+timeFormat()+".xlsx";
//		fileMap.put(path,targetPath);


		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_PRODUCT\\H_20171226_ASS_PRODUCT_0430.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\ASS_PRODUCT\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_ASS_PRODUCT_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\MAC\\H_20180108_MAC_2130.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\MAC\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_MAC_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Data3\\H_20180106_Base_Data3_1520.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Base_Data3\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Base_Data3_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_CODE\\H_20180109_Cto_CODE_1640.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_CODE\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Cto_CODE_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

		/*path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\REWORK\\H_20180105_Cto_REWORK_1600.xlsx";
		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\";
		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_REWORK_"+timeFormat()+".xlsx";
		fileMap.put(path,targetPath);*/

//		path="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Process\\H_20180116_Cto_Process_2000.xlsx";
//		targetPath="C:\\Users\\y15079\\Desktop\\shenjiangnan\\TestSample\\Cto_Process\\";
//		targetPath=targetPath+"Test_0_50000_H_"+dateFormat()+"_Cto_Process_"+timeFormat()+".xlsx";
//		fileMap.put(path,targetPath);

		Iterator<String> pathItr=fileMap.keySet().iterator();
		while (pathItr.hasNext()){
			path=pathItr.next();
			targetPath=fileMap.get(path);
			if (path.contains("ASS_ITEM")){
				excelAdd(path,targetPath,1,100000);
				/*copyFile(targetPath, 100,"ASS_ITEM");*/
			}
			if (path.contains("Base_Date(4)")){
				excelAdd(path,targetPath,2,50000);
//				copyFile(targetPath, 100,"Base_Date(4)");
			}
			if (path.contains("Po_sn")){
				excelAdd(path,targetPath,3,50000);
//				copyFile(targetPath, 100,"Po_sn");
			}
//			if (path.contains("Cto_Ship")){
//				excelAdd(path,targetPath,4,50000);
////				copyFile(targetPath, 100,"Cto_Ship");
//			}
//			if (path.contains("Cto_Process")){
//				excelAdd(path,targetPath,5,50000);
////				copyFile(targetPath, 100,"Cto_Process");
//			}
			if (path.contains("ASS_PRODUCT")){
				excelAdd(path,targetPath,6,50000);
//				copyFile(targetPath, 100,"ASS_PRODUCT");
			}
			if (path.contains("MAC")){
				excelAdd(path,targetPath,7,50000);
				copyFile(targetPath, 100,"MAC");
			}
			if (path.contains("Base_Data3")){
				excelAdd(path,targetPath,8,50000);
//				copyFile(targetPath, 100,"Base_Data3");
			}
			if (path.contains("Cto_CODE")){
				excelAdd(path,targetPath,9,50000);
//				copyFile(targetPath, 100,"Cto_CODE");
			}
			if (path.contains("Cto_REWORK")){
				excelAdd(path,targetPath,10, 50000);
//				copyFile(targetPath, 100,"Cto_REWORK");
			}
		}
	}

	public static void copyFile(String path,String targetPath) throws Exception{
		File file=new File(path);
		File[] fileArray=file.listFiles();
		File targetFile=new File(targetPath);
		if (!targetFile.exists()){
			targetFile.mkdirs();
		}
		byte[] b=new byte[1024];
		int n=0;
		for (int i=0;i<fileArray.length;i++){
			File file1=fileArray[i];
			FileInputStream fis=new FileInputStream(file);
			FileOutputStream fos=new FileOutputStream(targetPath+"\\"+file1.getName());
			while ((n=fis.read(b))!=-1){
				fos.write(b,0,n);
			}
			fis.close();
			fos.close();
		}
	}

	public static void copyFile(String path,int num,String type) throws Exception{
		File file=new File(path);
		String partentDir=file.getParentFile().getAbsolutePath();
		byte[] b=new byte[1024];
		int n=0;
		for (int i=1;i<=num;i++){
			String fileName="Test_"+i+"_100_H_"+dateFormat()+"_"+type+"_"+timeFormat()+".xlsx";
			FileInputStream fis=new FileInputStream(file);
			FileOutputStream fos=new FileOutputStream(partentDir+"/"+fileName);
			while ((n=fis.read(b))!=-1){
				fos.write(b,0,n);
			}
			fis.close();
			fos.close();
		}
	}

	public static String dateFormat(){
		SimpleDateFormat sf=new SimpleDateFormat("yyyyMMdd");
		return sf.format(new Date());
	}

	public static String timeFormat(){
		SimpleDateFormat sf=new SimpleDateFormat("HHmm");
		return sf.format(new Date());
	}

	public static void testBaseData(String path) throws Exception{
		File file=new File(path);
		InputStream is=new FileInputStream(file);
		XSSFWorkbook xssfWorkbook=new XSSFWorkbook(is);
		XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
	}

	public static void excelAdd(String path,String targetPath,int type,int num) throws Exception{
		File targetFile=new File(targetPath);
		if (!targetFile.exists()){
			targetFile.getParentFile().mkdirs();
		}
		File file=new File(path);
		InputStream is=new FileInputStream(file);
		XSSFWorkbook xssfWorkbook=new XSSFWorkbook(is);
		XSSFSheet xssfSheet=xssfWorkbook.getSheetAt(0);
		List<List<String>> rowList=new ArrayList<List<String>>();

		for (int rowNum=1; rowNum<=xssfSheet.getLastRowNum();rowNum++){
			XSSFRow xssfRow=xssfSheet.getRow(rowNum);
			int minColIx = xssfRow.getFirstCellNum();
			int maxColIx = xssfRow.getLastCellNum();
			List<String> cellList=new ArrayList<String>();
			boolean flag=false;

			for (int cellNum=minColIx;cellNum<maxColIx;cellNum++){
				XSSFCell xssfCell=xssfRow.getCell(cellNum);
				cellList.add(CommonUtils.toString(xssfCell));
				if (xssfCell != null && xssfCell.getCellType() != xssfCell.CELL_TYPE_BLANK){
					flag=true;
				}
			}
			if (flag){
				rowList.add(cellList);
			}
		}

		int loopNum=rowList.size();
		for (int count=loopNum+1;count<num;count=count+loopNum){
			for (int i=0;i<rowList.size();i++){
				List<String> cellList= rowList.get(i);
				XSSFRow xssfRow=xssfSheet.createRow(count+i);
				switch (type){
					case 1: AddExcel_Item(cellList,xssfRow,count+i+"");break;
					case 2: AddExcel_BaseBase1(cellList,xssfRow,count+i+""); break;
					case 3: AddExcel_PoSn(cellList,xssfRow,count+i+"");break;
					case 4: AddExcel_CotoShip(cellList,xssfRow,count+i+"");break;
					case 5: AddExcel_CotoProcess(cellList,xssfRow,count+i+"");break;
					case 6: AddExcel_CotoProduct(cellList,xssfRow,count+i+"");break;
					case 7: AddExcel_Mac(cellList,xssfRow,count+i+"");break;
					case 8: AddExcel_BaseData3(cellList,xssfRow,count+i+"");break;
					case 9: AddExcel_Code(cellList,xssfRow,count+i+"");break;
					case 10: AddExcel_Rework(cellList,xssfRow,count+i+"");break;
					default:break;
				}
			}
		}

		File file1=new File(targetPath);
		xssfWorkbook.write(new FileOutputStream(file1));
	}

	public static void AddExcel_Item(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j!=4){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}
		}
	}

	public static void AddExcel_BaseBase1(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==8){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_PoSn(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==6){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_CotoShip(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==4){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_CotoProcess(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==10){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_CotoProduct(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==11){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_Mac(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j==5){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}
		}
	}

	public static void AddExcel_BaseData3(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j!=6){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(countI);
			}
		}
	}

	public static void AddExcel_Code(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j!=5){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue("server"+countI);
			}
		}
	}

	public static void AddExcel_Rework(List<String> cellList , XSSFRow xssfRow, String countI){
		for (int j=0;j<cellList.size();j++){
			if (j!=22){
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue(cellList.get(j));
			}else {
				XSSFCell cell=xssfRow.createCell(j);
				cell.setCellValue("rework"+countI);
			}
		}
	}
}
