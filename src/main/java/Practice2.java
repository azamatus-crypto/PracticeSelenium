import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Practice2 {

    @Test(dataProvider = "driveTest")
    public void testCaseData(String name, String comunication, int numbers) {
        System.out.println(name + comunication + numbers);
    }


   @DataProvider(name = "getdata")
    public Object[]getData() throws IOException {
        DataFormatter dataFormatter = new DataFormatter();
        FileInputStream fileInputStream=new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Exel\\Exeldocuments2.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet=workbook.getSheetAt(2);
        int count=sheet.getPhysicalNumberOfRows();
        XSSFRow row=sheet.getRow(0);
        int columncount=row.getLastCellNum();
        Object [][]data=new Object[count-1][columncount];
        for (int i=0;i<count-1;i++){
            row = sheet.getRow(i + 1);
            for (int j = 0; j < columncount; i++) {//cyckle for cells in the row
                XSSFCell cell = row.getCell(j);
                data[i][j] = dataFormatter.formatCellValue(cell);
            }
        }
        return data;


    }
}
