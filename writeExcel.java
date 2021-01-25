package Excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;

public class writeExcel {
    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;
    public static int rowcount;

    //cach 3
    public static void copyExcelFile(String input, String output) throws IOException {

        //create file excel copy in java7
        File src = new File(input);
        File dest = new File(output);
        Files.copy(src.toPath(),dest.toPath());

    }
    public static void setExcelSheetWrite(String Pathin, String SheetName) throws IOException {
        try {
            //file input to Copy data from scenario file to result file,
            //vi moi vong lap, no tao file moi thi se chi luu ket qua cuoi nen de duong dan file out neu dung vong lap trong chuong trinh chinh
            FileInputStream ExcelFile = new FileInputStream(Pathin);

            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);

            //rowcount = ExcelWSheet.getLastRowNum() + 1; //the number of row
            //System.out.println("total rows is: " + rowcount);
        } catch (Exception e) {
            throw (e);
        }
    }


    public static void putCellData(int RowNum, int ColNum, String value) throws IOException {
        Cell = ExcelWSheet.getRow(RowNum).createCell(ColNum);
        Cell.setCellValue(value);

        //file output for loginDEMOQAexcel.run.loginTest
        File out = new File("src/test/resources/result_loginDataDEMOQA.xlsx");


        FileOutputStream fout = new FileOutputStream(out);
        ExcelWBook.write(fout);
        ExcelWBook.close();

    }
}

/*

    public static void setExcelFileWrite(String Path, String SheetName) throws Exception {
        try{
            // Open the Excel file
            FileInputStream ExcelFile = new FileInputStream(Path);
            // Access the required test data sheet
            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);

            //System.out.println("total rows is: " + config.getCellData(1,1));
            XSSFSheet Sheet1 = ExcelWBook.getSheet("Sheet1");

            rowcount = Sheet1.getLastRowNum() + 1; //the number of row

        } catch (Exception e) {
            throw (e);
        }

    }

    public static String putCellData(int RowNum, int ColNum)throws Exception {

        //Copy data from scenario file to result file

        for(int i = 1; i < rowcount-1; i++) {
            sheet1.getRow(i).createCell(5).setCellValue("pass");

        }
        try {
            Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);

            String CellData = Cell.setCellValue();
            return CellData;
        } catch (Exception e) {
            return "";
        }

        Sheet1.getRow(0).createCell(5).setCellValue("Result");
        for(int i = 1; i < rowcount-1; i++) {

            sheet1.getRow(i).createCell(5).setCellValue("pass");

        }
        sheet1.getRow(rowcount-1).createCell(5).setCellValue("fail");

        FileOutputStream fout = new FileOutputStream(src);
    }
}
*/


/*    //cach 2


    public static void main(String[] args) throws IOException {
        //file output
        File src = new File("src/test/resources/result_loginDataDEMOQA.xlsx");

        FileInputStream fin = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fin);
        XSSFSheet sheet1 = wb.getSheetAt(0);
        int rowcount = sheet1.getLastRowNum() + 1; //the number of row
        System.out.println("total rows is: " + rowcount);

        sheet1.getRow(0).createCell(5).setCellValue("Result");
        for (int i = 1; i < rowcount - 1; i++) {

            sheet1.getRow(i).createCell(5).setCellValue("pass");

        }
        sheet1.getRow(rowcount - 1).createCell(5).setCellValue("fail");

        FileOutputStream fout = new FileOutputStream(src);
        wb.write(fout);
        wb.close();
    }*/





