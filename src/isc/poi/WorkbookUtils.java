package isc.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.*;

public class WorkbookUtils
{
    Workbook workbook = null;

    public String init(String filename)
    {
        String result = "";
        File file = new File(filename);

        try
        {
            workbook = WorkbookFactory.create(file);
        }
        catch (Exception e)
        {
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            String exceptionAsString = sw.toString();
            result = e.getLocalizedMessage() + " " + exceptionAsString;
        }

        return result;
    }

    public String finishFillAndCreateResultFile(String outFilename)
    {
        String result = "";
        try
        {
            //workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            workbook.setForceFormulaRecalculation(true);
            FileOutputStream out = new FileOutputStream(outFilename);
            workbook.write(out);
            out.close();
        }
        catch (Exception e)
        {
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            String exceptionAsString = sw.toString();
            result = e.getLocalizedMessage() + " " + exceptionAsString;
        }

        return result;
    }

    public String fillBookSheet(int sheetNumber, String[] params)
    {
        String result = "";

        try
        {
            Sheet sheet = workbook.getSheetAt(sheetNumber);

            for (String param: params)
            {
                String[] paramArray = param.split("\u0001");
                String address = paramArray[0];
                String value;

                if (paramArray.length > 1)
                {
                    value = paramArray[1];
                } else {
                    value = "";
                }

                CellReference cellReference = new CellReference(address);
                Row row = sheet.getRow(cellReference.getRow());

                if (row == null) {
                    row = sheet.createRow(cellReference.getRow());
                }

                Cell cell = row.getCell(cellReference.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (value.matches("\\d+(\\.\\d+)?"))
                {
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(Double.parseDouble(value));
                }else
                {
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(value);
                }
            }
        } catch (Exception e) {
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            String exceptionAsString = sw.toString();

            result = e.toString() + " " + exceptionAsString;
        }
        return result;
    }
}
