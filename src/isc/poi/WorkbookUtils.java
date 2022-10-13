package isc.poi;

import com.pnuema.java.barcode.Barcode;
import com.pnuema.java.barcode.EncodingType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Objects;

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
                String type;

                if (paramArray.length > 1)
                {
                    value = paramArray[1];
                } else {
                    value = "";
                }

                if (paramArray.length > 2)
                {
                    type = paramArray[2].toUpperCase();
                } else {
                    type = "";
                }

                CellReference cellReference = new CellReference(address);
                Row row = sheet.getRow(cellReference.getRow());

                if (row == null) {
                    row = sheet.createRow(cellReference.getRow());
                }

                Cell cell = row.getCell(cellReference.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                if (Objects.equals(type, "NUMERIC"))
                {
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(Double.parseDouble(value));
                }
                else if (Objects.equals(type, "BARCODE"))
                {
                    byte[] data = createBarCode(value);
                    inputImage(sheet, cell, data);
                }
                else
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

    public void inputImage(Sheet sheet, Cell cell, byte[] data) throws IOException
    {
        int inputImagePictureID1 = workbook.addPicture(data, Workbook.PICTURE_TYPE_PNG);

        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        XSSFClientAnchor ironManAnchor = new XSSFClientAnchor();

        int row = cell.getRowIndex();
        int col = cell.getColumnIndex();

        ironManAnchor.setCol1(col);   // Sets the column (0 based) of the first cell.
        ironManAnchor.setCol2(col+1); // Sets the column (0 based) of the Second cell.
        ironManAnchor.setRow1(row);   // Sets the row (0 based) of the first cell.
        ironManAnchor.setRow2(row+1); // Sets the row (0 based) of the Second cell.

        drawing.createPicture(ironManAnchor, inputImagePictureID1);
    }

    public byte[] createBarCode(String stringToEncode) throws IOException
    {
        //https://github.com/barnhill/barcode-java
        Barcode barcode = new Barcode();
        BufferedImage img = (BufferedImage) barcode.encode(EncodingType.UPCA, stringToEncode);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(img, "png", bos );

        return bos.toByteArray();
    }
}
