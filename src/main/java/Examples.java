import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class Examples {


    public ArrayList<String> getInfoFromExel(String name) throws IOException {
        int countOfCell = 0;
        int totallcells = 0;
        ArrayList<String> arrayList = new ArrayList<>();
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Exel\\Exeldocuments2.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        int countOfSheet = workbook.getNumberOfSheets();
        for (int i = 0; i < countOfSheet; i++) {
            if (workbook.getSheetName(i).equals("Sheet2")) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                Iterator<Row> rowIterator = sheet.iterator();
                rowIterator.next();
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                int k = 0;
                int column = 0;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getStringCellValue().equals("User Name")) {
                        column = k;
                    }
                    k++;
                }
                System.out.println(column);
                while (rowIterator.hasNext()) {
                    Row row2 = rowIterator.next();
                    if (row2.getCell(column).getStringCellValue().equalsIgnoreCase(name)) {
                        Iterator<Cell> cellIterator2 = row2.cellIterator();
                        while (cellIterator2.hasNext()) {
                            Cell c= cellIterator2.next();
                            if (c.getCellType() == CellType.STRING) {
                                arrayList.add(c.getStringCellValue());
                            } else {
                                arrayList.add(NumberToTextConverter.toText(c.getNumericCellValue()));
                            }
                        }
                    }
                }
            }


        }
        return arrayList;
    }
}






