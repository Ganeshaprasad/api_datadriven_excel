import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;

import java.io.IOException;
import java.util.Iterator;

public class dataDriven {
    /**
     * Identify the test case column by scanning the entire 1st row
     * Once colunm identified then scan entire test case to identify purchase test case
     * After grab purchase row pull all data
     *
     * @throws IOException
     */
    @Test
    public static void dataDrivenFromExcel() throws IOException {
        //get java representative of physical file
        FileInputStream fis = new FileInputStream("C://Users//prasa//Downloads//dataDriven.xlsx");
        //get access to file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        // get access to sheet
        int sheets = workbook.getNumberOfSheets();
        for (int i = 0; i < sheets; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
                XSSFSheet sheet = workbook.getSheetAt(i);

                //acces to specific row
                Iterator<Row> rows = sheet.iterator();
                Row firstRow = rows.next();
                // access to column
                Iterator<Cell> cell = firstRow.cellIterator();

                int k = 0;
                int column = 0;
                while (cell.hasNext())// it will check next cell there or not
                {
                    Cell value = cell.next();// moved to nxt cell


                    if (value.getStringCellValue().equalsIgnoreCase("Data2")) {
                        column = k;
                    }
                    k++;
                }
                System.out.println(column);
            }
        }
    }
}
