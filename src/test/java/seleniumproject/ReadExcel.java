package seleniumproject;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcel
{
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("C:\\Users\\gayat\\OneDrive\\Desktop\\excel\\TestData.xlsx");
    XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet("Sheet1");
        int rowcount = sheet.getPhysicalNumberOfRows();
        System.out.println("No of rows: "+rowcount);
        int colcount = sheet.getRow(0).getPhysicalNumberOfCells();
        System.out.println("No of columns: "+colcount);
        for(int i=0;i<rowcount;i++) {
            for(int j=0;j<colcount;j++) {
                String value =sheet.getRow(i).getCell(j).toString();
                System.out.print(value+" | ");
            }
            System.out.println();




        }

        wb.close();
        fis.close();
    }}










