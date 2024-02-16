import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {
    public static ArrayList<String> getDataFromExel(String name) throws IOException {
        ArrayList<String>str=new ArrayList<>();
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\Exel\\Exeldocuments2.xlsx");
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        int totlalSheets = xssfWorkbook.getNumberOfSheets();
        for (int i = 0; i < totlalSheets; i++) {
            if (xssfWorkbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
                XSSFSheet xs = xssfWorkbook.getSheetAt(i);
                Iterator<Row> it = xs.iterator();
                Row r = it.next();
                Iterator<Cell>ce=r.cellIterator();
                int k=0;
                int column=0;
                while (ce.hasNext()){
                    Cell s=ce.next();
                    if(s.getStringCellValue().equals("Test case")){
                        column=k;
                    }
                    k++;
                }
                System.out.println(column);
                while (it.hasNext()){
                    Row row=it.next();
                    if (row.getCell(k).getStringCellValue().equals(name)){
                        Iterator<Cell>cellIterator= row.cellIterator();
                        while (cellIterator.hasNext()){
                            Cell atr=cellIterator.next();
                            if (atr.getCellType()== CellType.STRING){
                                str.add(cellIterator.next().getStringCellValue());
                            }else {
                                str.add(NumberToTextConverter.toText(cellIterator.next().getNumericCellValue()));
                            }

                        }
                    }
                }

            }

        }
        return str;
    }
    public static void main(String[] args) throws IOException {
      getDataFromExel("Purchase");

    }
}
