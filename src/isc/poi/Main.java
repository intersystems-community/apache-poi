package isc.poi;

import com.intersys.jdbc.CacheListBuilder;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;

public class Main {

    private static final char separator = (char)1;

    public static String[] getBookFromStream(byte[] stream) {
        String[] result = new String[1];


        ArrayList<String> bookList = new ArrayList<String>();
        CacheListBuilder bookInfo = new CacheListBuilder("UTF8");

        Workbook workbook = loadBook(new ByteArrayInputStream(stream), result);

        if (result[0]==null) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                CacheListBuilder sheetInfo = new CacheListBuilder("UTF8");

                Sheet sheet = sheetIterator.next();
                ArrayList<String> rowList = getSheetInternal(sheet);
                bookList.addAll(rowList);

                try {
                    sheetInfo.set(rowList.size());
                    sheetInfo.set(sheet.getSheetName().getBytes());
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

    private static Workbook loadBook(String filename, String[] result) {
        Workbook workbook = null;
        File file = new File(filename);

        try {
            workbook = WorkbookFactory.create(file);
        } catch (Exception e) {
            // IOException InvalidFormatException
            result[0] = e.toString();
        }

        return workbook;
    }

    private static Workbook loadBook(InputStream stream, String[] result) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(stream);
        } catch (Exception e) {
            // IOException InvalidFormatException
            result[0] = e.toString();
        }

        return workbook;
    }

    public static String[] getBook(String filename) {
        String[] result = new String[1];
        ArrayList<String> bookList = new ArrayList<String>();
        CacheListBuilder bookInfo = new CacheListBuilder("UTF8");
        Workbook workbook = loadBook(filename, result);

        if (result[0]==null) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                CacheListBuilder sheetInfo = new CacheListBuilder("UTF8");

                Sheet sheet = sheetIterator.next();
                ArrayList<String> rowList = getSheetInternal(sheet);
                bookList.addAll(rowList);

                try {
                    sheetInfo.set(rowList.size());
                    sheetInfo.set(sheet.getSheetName().getBytes());
                    bookInfo.set(new String(sheetInfo.getData()));
                } catch (SQLException e) {
                    // does not seem to be throwable
                }
            }
            bookList.add(new String(bookInfo.getData()));
            result = bookList.toArray(new String[bookList.size()]);
        }
        try {
            workbook.close();
        } catch (IOException e) {
            result = new String[1];
            result[0] = e.toString();
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

                try {
                    workbook.close();
                } catch (IOException e) {
                    result = new String[1];
                    result[0] = e.toString();
                }
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
            String list= "";
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
                            case ERROR:
                                value = "";
                                break;
                        }
                        break;
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


                if (value instanceof String) {
                    list += (String)value;
                } else if (value instanceof Double) {
                    Double doubleValue = (Double)value;
                    if (doubleValue == Math.floor(doubleValue) && !Double.isInfinite(doubleValue)) {
                        list += doubleValue.longValue();
                    } else {
                        list += value.toString();
                    }
                } else {
                    list += value.toString();
                }
                list += separator;
            }
            // remove last separator
            rowList.add(list.substring(0, list.length() - 1));
        }
        return rowList;
    }

    public static int getSheetCount(String filename) throws Exception {
        File file = new File(filename);
        Workbook workbook = WorkbookFactory.create(file);
        return workbook.getNumberOfSheets();
    }

}
