package org.example;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.io.*;

import java.util.List;
import java.util.concurrent.Callable;
public abstract class excelBaseClass {
    protected String filepath = null;
    protected File file = null;
}

class ExcelReader extends excelBaseClass
{
    FileInputStream fis = null;
    XSSFWorkbook wb = null;
    ExcelReader(String Filepath) throws IOException {
        this.filepath = Filepath;

        File filer = new File(this.filepath);
        if(filer.exists())
        {
            this.file = filer;
            this.fis = new FileInputStream(this.file);
            this.wb = new XSSFWorkbook(this.fis);
            this.fis.close();
        } else {
            this.file = null;
        }
    }


    public void DisplayFile()
    {
        if(this.file == null)
        {
            throw new RuntimeException("FIle does not exist");
        }

        FormulaEvaluator formulaEval = this.wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet=wb.getSheetAt(0);
        for(Row row:sheet)
        {
            for(Cell cell: row)
            {
                switch(formulaEval.evaluateInCell(cell).getCellType())
                {
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        break;

                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        System.out.printf("%4s", cell.getCellFormula());
                        break;
                }
            }
            System.out.println();
        }
    }
}

class ExcelWriter extends excelBaseClass
{



    public static void editExcelFile(Object[][] bookData, File filer)  {
        if(!filer.exists())
        {
            Workbook wb = new XSSFWorkbook();
            try {
                FileOutputStream temp = new FileOutputStream(filer.getAbsolutePath());
                wb.createSheet();
                wb.write(temp);
            } catch (FileNotFoundException e) {
                System.err.println(e.getMessage());
            } catch (IOException e) {
                System.err.println(e.getMessage());
            }
        }
        try {
            FileInputStream fip = new FileInputStream((filer));
            XSSFWorkbook wb = new XSSFWorkbook(fip);
            Sheet sheet = wb.createSheet("Alladi_Sheet");
            int rowCount = sheet.getLastRowNum();
            int colCount = 0;

            for(Object[]abk : bookData)
            {
                int colnum = 0;
                Row newrow = sheet.createRow(++rowCount);

                for(Object dat : abk)
                {
                    Cell cell = newrow.createCell(colnum);
                    if(dat instanceof String)
                    {
                        cell.setCellValue((String)dat);
                    } else if(dat instanceof Integer)
                    {
                        cell.setCellValue((Integer)dat);
                    }

                    ++colnum;
                }
            }
            FileOutputStream fos = new FileOutputStream(filer);
            wb.write(fos);
            wb.close();
            fos.close();

        } catch (FileNotFoundException e) {
            System.err.println(e.getMessage());
        } catch (IOException e) {
            System.err.println(e.getMessage());
        }




    }

}
