package seleniumproject;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteToexcel {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("C:\\Users\\gayat\\OneDrive\\Desktop\\excel\\TestData2.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);

        wb.createSheet().createRow(0).createCell(0).setCellValue("Gayatri Manjani");

        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\gayat\\OneDrive\\Desktop\\excel\\TestData2.xlsx");
        wb.write(outputStream);
        wb.close();

        wb.close();
    }

}
