package isc.poi;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.util.CellReference;

public class Main {

    private static final char separator = (char)1;

    public static String[] getBookFromStream(byte[] stream) {
        String[] result = new String[1];


        ArrayList<String> bookList = new ArrayList<String>();
        String bookInfo = "";

        Workbook workbook = loadBook(new ByteArrayInputStream(stream), result);

        if (result[0]==null) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                ArrayList<String> rowList = getSheetInternal(sheet);
                bookList.addAll(rowList);

                bookInfo += String.valueOf(rowList.size()) + (char)1;
                bookInfo += sheet.getSheetName() + (char)1 + (char)1;
            }
            bookInfo = bookInfo.substring(0, bookInfo.length() - 2);
            bookList.add(bookInfo);
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
        String bookInfo = "";
        Workbook workbook = loadBook(filename, result);

        if (result[0]==null) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while(sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                ArrayList<String> rowList = getSheetInternal(sheet);
                bookList.addAll(rowList);

                bookInfo += String.valueOf(rowList.size()) + (char)1;
                bookInfo += sheet.getSheetName() + (char)1 + (char)1;
            }

            bookInfo = bookInfo.substring(0, bookInfo.length() - 2);
            bookList.add(bookInfo);
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

    public static String fillBook(String filename, String outFilename,  int sheetNumber, String[] params){
        String result = "";
        try {
            File file = new File(filename);

            Workbook workbook = null;
            workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(sheetNumber);
            for (String param: params)
            {
                String[] paramArray = param.split("\u0001");
                String address = paramArray[0];
                String value = paramArray[1];

                CellReference cellReference = new CellReference(address);
                Row row = sheet.getRow(cellReference.getRow());

                if (row == null) {
                    row = sheet.createRow(cellReference.getRow());
                }

                Cell cell = row.getCell(cellReference.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(value);
            }

            FileOutputStream out = new FileOutputStream(outFilename);
            workbook.write(out);
            out.close();

        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            String exceptionAsString = sw.toString();

            result = e.toString() + " " + exceptionAsString;
        }
        //File file = new File(filename);
        //Workbook workbook = WorkbookFactory.create(file);
        return result;
    }

}
