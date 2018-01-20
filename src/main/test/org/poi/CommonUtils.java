package org.poi;

import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.LocaleUtil;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author y15079
 * @create 2017-12-09 13:35
 * @desc
 **/
public class CommonUtils {

	public static String getDateToString(Date date){
		DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return sdf.format(date);
	}

	public static String toString(Cell cell){
			switch(cell.getCellTypeEnum()) {
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", LocaleUtil.getUserLocale());
						sdf.setTimeZone(LocaleUtil.getUserTimeZone());
						return sdf.format(cell.getDateCellValue());
					}
					return Double.toString(cell.getNumericCellValue());
				case STRING:
					return cell.getRichStringCellValue().toString();
				case FORMULA:
					return cell.getCellFormula();
				case BLANK:
					return "";
				case BOOLEAN:
					return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
				case ERROR:
					return ErrorEval.getText(cell.getErrorCellValue());
				default:
					return "Unknown Cell Type: " + cell.getCellTypeEnum();
			}
	}
}
