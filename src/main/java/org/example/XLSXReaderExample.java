package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlOptions;

public class XLSXReaderExample
    {
        public static void main(String[] args)
        {
            try
            {
/*                ClassLoader classloader =
                        org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
                URL res = classloader.getResource(
                        "org/apache/poi/poifs/filesystem/POIFSFileSystem.class");*/

                ClassLoader classloader =
                        org.apache.xmlbeans.XmlOptions.class.getClassLoader();
                URL res = classloader.getResource(
                        "org/apache/xmlbeans/XmlOptions.class");



                String path = res.getPath();
                System.out.println("Core POI came from " + path);

                File file = new File("/home/dev/dev/ReadExcel/Book1.xlsx");   //creating a new file instance
                FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
                //creating Workbook instance that refers to .xlsx file
                XSSFWorkbook wb = new XSSFWorkbook(fis); // error in this line
                XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
                Iterator<Row> itr = sheet.iterator();    //iterating over excel file
                while (itr.hasNext())
                {
                    Row row = itr.next();
                    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                    while (cellIterator.hasNext())
                    {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType())
                        {
                            case STRING:    //field that represents string cell type
                                System.out.print(cell.getStringCellValue() + "\t\t\t");
                                break;
                            case NUMERIC:    //field that represents number cell type
                                System.out.print(cell.getNumericCellValue() + "\t\t\t");
                                break;
                            default:
                        }
                    }
                    System.out.println("");
                }
            }
            catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }
