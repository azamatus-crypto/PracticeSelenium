import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Practice {
    DataFormatter dataFormatter = new DataFormatter();

    @Test(dataProvider = "driveTest")
    public void testCaseData(String name, String comunication, int numbers) {
        System.out.println(name + comunication + numbers);
    }

    @DataProvider(name = "driveTest")
    public Object[] getData() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Exel\\Exeldocuments2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(2);
        int rowCount = sheet.getPhysicalNumberOfRows();//got physical rows
        XSSFRow row = sheet.getRow(0);
        int columncount = row.getLastCellNum();
        Object[][] data = new Object[rowCount - 1][columncount];//-1 becouse dont wanna read info on the top
        for (int i = 0; i < rowCount - 1; i++) {//cycle for orws
            row = sheet.getRow(i + 1);
            for (int j = 0; j < columncount; i++) {//cyckle for cells in the row
                XSSFCell cell = row.getCell(j);
                data[i][j] = dataFormatter.formatCellValue(cell);
            }
        }

      return data;
    }
}
