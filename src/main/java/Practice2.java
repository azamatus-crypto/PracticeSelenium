import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Practice2 {




    public Object[]getData() throws IOException {
        FileInputStream fileInputStream=new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Exel\\Exeldocuments2.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet=workbook.getSheetAt(2);
        int count=sheet.getPhysicalNumberOfRows();
        XSSFRow row=sheet.getRow(0);
        int columncount=row.getLastCellNum();
        Object [][]data=new Object[count-1][columncount];

        return data;


    }
}
