package com.etohtm;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import org.w3c.dom.Document;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        System.out.println("Hello World!\nExcel to HTML");

        String excelFilePath = "resources\\excel.xls";

         try { 
        InputStream inp = new FileInputStream(excelFilePath);

       HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
        ExcelExtractor extractor = new ExcelExtractor(wb);

        extractor.setFormulasNotResults(false);
        extractor.setIncludeSheetNames(false);
        String text = extractor.getText();
       
        
          // Iterate over sheets
            for (Sheet sheet : wb) {
                sheet.setForceFormulaRecalculation(true);
                StringBuilder html = new StringBuilder();
                html.append("<!DOCTYPE html>\n");
                html.append("<html>\n");
                html.append("<head></head>\n");
                html.append("<body>\n");
                html.append("<table><thead><tr>");

                // Header row
                Row headerRow = sheet.getRow(0);
                for (Cell cell : headerRow) {
                    
                    html.append("<th>").append(getCellHtmlValue(cell)).append("</th>");
                }
                html.append("</tr></thead><tbody>");

                // Data rows
                for (int rowNum = 1; rowNum < sheet.getLastRowNum() + 1; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (row != null ) {                    
                        html.append("<tr>");
                        
                        for (Cell cell : row) {
                            
                            html.append("<td>").append(getCellHtmlValue(cell)).append("</td>");
                        }
                        html.append("</tr>");
                        }
                    }

                    html.append("</tbody></table>");
                    html.append("</body>");
                    html.append("</html>");

                    // Write the HTML to a file or print to console
                  FileWriter fr = new FileWriter("etoh.html");
                   fr.append(html);
                   fr.close();
                    System.out.println(html.toString());
                }
            } catch (IOException e) {
                e.printStackTrace();
            }  

    }
    private static String getCellHtmlValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                switch(cell.getCachedFormulaResultType()){
                   case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue());
                }
                
            default:
                return "";
        }
    }
}
