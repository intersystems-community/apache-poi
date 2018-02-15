package isc.poi;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Iterator;
import java.io.File;

import static org.apache.poi.ss.usermodel.CellType.*;

public class Main {

    public static String ROWSEPARATOR = "\t\t\t";

    public static void main(String[] args) {
        try {
            Test1();
        } catch (Exception ex) {
        }
    }

    public static String[] Test1 () throws Exception{
        ArrayList<String> list = new ArrayList<String>();

        File file = GetFile();
        Workbook workbook = WorkbookFactory.create(file);
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while(sheetIterator.hasNext()){
            Sheet sheet = sheetIterator.next();
            String name  = sheet.getSheetName();
            String value = null;

            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                Row row = (Row) rows.next();

                for(int i=0; i<row.getLastCellNum(); i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if (cell.getCellTypeEnum() == FORMULA) {
                        switch(cell.getCachedFormulaResultTypeEnum()) {
                            case NUMERIC:
                                value = String.valueOf(cell.getNumericCellValue());
                                break;
                            case STRING:
                                value = cell.getRichStringCellValue().getString();
                                break;
                        }
                    } else {
                        value = cell.toString();
                    }
                    list.add(value);
                    ///System.out.print("'" + cell.toString() + "'"+" ");
                }
                list.add(ROWSEPARATOR);
                System.out.println();
            }


            /*for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString()+" ");
                    //int i=1;
                }
                System.out.println();
            }*/
        }
        String[] result = list.toArray(new String[list.size()]);
        return result;
    }

    public static File GetFile () {
        File file = new File("D:\\Cache\\POI\\Книга1.xlsx");

        return file;
    }
}
