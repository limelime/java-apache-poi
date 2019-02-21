package net.openwritings.xngo.tutorial.poi;

/**
 * Example showing how to write and read Excel file(i.e *.xls or *.xlsx).
 * JAR files needed:
 *    poi-*.jar
 *    poi-ooxml-*.jar
 * If you only need to handle Excel 2007 OOXML (.xlsx) file format, then you can use XSSF* classes.
 * If you only need to handle Excel 97-2003(.xls) file format, then you can use HSSF* classes.
 * @author Xuan Ngo
 */
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
 
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.WorkbookFactory; // This is included in poi-ooxml-*.jar
import org.apache.poi.ss.usermodel.Workbook;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 
public class PoiExample{
 
    public static void main(String[] args){
        try{
            // Create an Excel file.
            //***********************************
            Workbook writeWorkbook = new HSSFWorkbook();
            Sheet sheet1 = writeWorkbook.createSheet("new sheet");
 
            Row row1 = sheet1.createRow(0);
            Cell cell1 = row1.createCell(0);
            cell1.setCellValue("Xuan");
 
            // Write workbook to a file.
            FileOutputStream fileOut = new FileOutputStream("new_workbook.xls");
            writeWorkbook.write(fileOut);
            fileOut.close();
            writeWorkbook.close();
 
            // Read an Excel file.
            //***********************************
 
            // WorkbookFactory create the appropriate kind of Workbook (be it HSSFWorkbook or XSSFWorkbook), 
            //	by auto-detecting from the supplied input.
            Workbook readWorkbook = WorkbookFactory.create(new FileInputStream("new_workbook.xls") );
 
            // Get the first sheet.
            Sheet sheet = readWorkbook.getSheetAt(0);
 
            // Get the first cell.
            Row row = sheet.getRow(0);
            Cell cell = row.getCell(0);
 
            // Show what is being read.
            System.out.println("Read cell(0,0)="+cell.toString());
 
            
            // Close the workbook.
            readWorkbook.close();
        }
        catch(FileNotFoundException e){
            System.out.println(e);
        }
        catch(IOException e){
            System.out.println(e);
        }
 
    }
}
