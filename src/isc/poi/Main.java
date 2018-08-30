package isc.poi;

import com.intersys.jdbc.CacheListBuilder;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row;

import java.io.IOException;
import java.math.BigDecimal;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.io.File;

import static org.apache.poi.ss.usermodel.CellType.*;

public class Main {

    public static String[] getBook(String filename) {
        File file = new File(filename);

        Workbook workbook = null;
        String[] result = new String[1];
        ArrayList<String> bookList = new ArrayList<String>();
        CacheListBuilder bookInfo = new CacheListBuilder("UTF8");

        try {
            workbook = WorkbookFactory.create(file);
        } catch (Exception e) {
            // IOException InvalidFormatException
            result[0] = e.toString();
        }

        if (result[0]==null) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                CacheListBuilder sheetInfo = new CacheListBuilder("UTF8");

                Sheet sheet = sheetIterator.next();
                ArrayList<String> rowList = getSheetInternal(sheet);
                bookList.addAll(rowList);

                try {
                    sheetInfo.set(rowList.size());
                    sheetInfo.set(sheet.getSheetName());
                    bookInfo.set(new String(sheetInfo.getData()));
                } catch (SQLException e) {
                    // does not seem to be throwable
                }
            }
            bookList.add(new String(bookInfo.getData()));
            result = bookList.toArray(new String[bookList.size()]);
        }
        return result;
    }

    /// Iterators - do not skip empty rows?
    /// https://stackoverflow.com/questions/30519539/apache-poi-skips-rows-that-have-never-been-updated
    public static String[] getSheet(String filename, int sheetNumber) {
        File file = new File(filename);

        Workbook workbook = null;
        String[] result = new String[1];

        try {
            workbook = WorkbookFactory.create(file);
        } catch (Exception e) {
            // IOException InvalidFormatException
            result[0] = e.toString();
        }

        if (result[0]==null) {
            Sheet sheet = null;
            try {
                sheet = workbook.getSheetAt(sheetNumber);
            } catch (Exception e) {
                result[0] = e.toString();
            }

            if (result[0]==null) {
                ArrayList<String> rowList = getSheetInternal(sheet);
                result = rowList.toArray(new String[rowList.size()]);
            }
        }
        return result;
    }

    /// Pass ArrayList here?
    private static ArrayList<String> getSheetInternal(Sheet sheet) {

        Object value = null;
        ArrayList<String> rowList = new ArrayList<String>();
        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            CacheListBuilder list = new CacheListBuilder("UTF8");
            Row row = (Row) rows.next();

            for (int i = 0; i < row.getLastCellNum(); i++) {
                value = "";
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                switch (cell.getCellTypeEnum()) {
                    case FORMULA:
                        switch (cell.getCachedFormulaResultTypeEnum()) {
                            case NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    value = new java.sql.Date(cell.getDateCellValue().getTime());
                                } else {
                                    value = cell.getNumericCellValue();
                                }
                                break;
                            case STRING:
                                value = cell.getRichStringCellValue();
                                break;
                        }
                    case NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            value = new java.sql.Date(cell.getDateCellValue().getTime());
                        } else {
                            value = cell.getNumericCellValue();
                        }
                        break;
                    case STRING:
                        value = cell.getRichStringCellValue().getString();
                        break;
                    case _NONE:
                    case BLANK:
                    case ERROR:
                        // not a null
                        value = "";
                        break;
                    case BOOLEAN:
                        value = cell.getBooleanCellValue();
                }



                try {

                    if (value instanceof String){
                        // TODO - other solutions for 2byte strings? lines 195-200 contain error
                        list.set(((String)value).getBytes());
                    } else if (value instanceof Double) {
                        Double doubleValue = (double)value;
                        if ((doubleValue == Math.floor(doubleValue) && !Double.isInfinite(doubleValue))) {
                            // integer type
                            list.set(doubleValue.longValue());
                        } else {
                            // TODO Set them as numbers.
                            // BigDouble does not work.
                            list.set(value.toString());
                        }
                    } else {
                        list.setObject(value);
                    }

                } catch (SQLException ex){
                    // does not seem to be throwable
                }
            }
            rowList.add(new String(list.getData()));
        }

        return rowList;
    }

    public static int getSheetCount(String filename) throws Exception {
        File file = new File(filename);
        Workbook workbook = WorkbookFactory.create(file);
        return workbook.getNumberOfSheets();
    }

}
