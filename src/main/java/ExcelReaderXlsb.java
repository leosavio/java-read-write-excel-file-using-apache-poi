import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.binary.XSSFBParseException;
import org.apache.poi.xssf.binary.XSSFBParser;
import org.apache.poi.xssf.extractor.XSSFBEventBasedExcelExtractor;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class ExcelReaderXlsb {
    public static final String SAMPLE_XLS_FILE_PATH = "./sample-xls-file.xls";
    public static final String SAMPLE_XLSX_FILE_PATH = "./sample-xlsx-file.xlsx";
    public static final String SAMPLE_XLSB_FILE_PATH = "./sample-xls-file.xlsb";

    public static void main(String[] args) throws IOException, InvalidFormatException {
    	
    	XSSFBEventBasedExcelExtractor ext = null;
    	try {
    	    ext = new XSSFBEventBasedExcelExtractor(SAMPLE_XLSB_FILE_PATH);
    	    System.out.println(ext.getText());
    	} catch (Exception ex) {
    	    System.out.println(ex.getMessage());
    	}
    	

    }

    private static void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");
    }
}
