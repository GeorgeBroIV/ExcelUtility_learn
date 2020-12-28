package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.IOException;
import org.junit.jupiter.api.Test;

public class ExcelUtils {

    public static void main(String[] args) throws IOException {
        try {
            getRowCount();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void getRowCount() throws IOException {

        try {
            String excelPath = "./data/";
            String excelFileName = "TestData.xlsx";
            String excelFile = excelPath + excelFileName;
            String excelSheet = "Sheet1";

            XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
            XSSFSheet sheet = workbook.getSheet(excelSheet);
            int rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("Number of rows in Excel workbook '" + excelFileName + "' sheet '" + excelSheet +"': " + rowCount);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
