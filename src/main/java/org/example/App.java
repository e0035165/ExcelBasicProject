package org.example;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException {

        ExcelReader reader = new ExcelReader("C:\\Users\\ACER\\Documents\\Alladi Satya Chandra Bharathi -Uob Timesheet -September.xlsx");
        reader.DisplayFile();

        Object[][] bookData = {
                {"The Passionate Programmer", "Chad Fowler", 16},
                {"Software Craftmanship", "Pete McBreen", 26},
                {"The Art of Agile Development", "James Shore", 32},
                {"Continuous Delivery", "Jez Humble", 41},
                {"Fantastic four age of ultron", "James Bron", 12}
        };

        //ExcelWriter writer = new ExcelWriter("C:\\Users\\ACER\\Documents\\Test.xlsx");
        File filer = new File("C:\\Users\\ACER\\Documents\\Books_Datafile.xlsx");
        ExcelWriter.editExcelFile(bookData, filer);
    }
}
