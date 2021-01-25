package Excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class readExcel {

    private static XSSFSheet ExcelWSheet;
    private static XSSFWorkbook ExcelWBook;
    private static XSSFCell Cell;
    private static XSSFRow Row;
    public static int rowcount;
    public static void setExcelFileRead(String Path, String SheetName) throws Exception {
        try {
            // Open the Excel file
            FileInputStream ExcelFile = new FileInputStream(Path);
            // Access the required test data sheet
            ExcelWBook = new XSSFWorkbook(ExcelFile);
            ExcelWSheet = ExcelWBook.getSheet(SheetName);

            //System.out.println("total rows is: " + config.getCellData(1,1));
            //Sheet need to read
            XSSFSheet Sheetname = ExcelWBook.getSheet(SheetName);

            //rowcount = Sheetname.getLastRowNum() + 1; //the number of row
        } catch (Exception e) {
            throw (e);
        }
    }

    public static String getCellData(int RowNum, int ColNum) throws IOException {
        try {
            Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
            String CellData = Cell.getStringCellValue();
            return CellData;
        } catch (Exception e) {
            return "";
        }
    }


}
/*
    public static void main(String[] args) throws Exception {
        readExcel config = new readExcel();
        config.setExcelFile("src/test/resources/loginDataDEMOQA.xlsx","sheet1");

        //System.out.println("total rows is: " + config.getCellData(1,1));
        XSSFSheet Sheet1 = ExcelWBook.getSheet("Sheet1");

        int rowcount = Sheet1.getLastRowNum() + 1; //the number of row
        //System.out.println("total rows is: " + rowcount);

        for (int i = 0; i < rowcount; i++)
            for (int j = 0;j<7 ;j++){
                String data = getCellData(i,j);
                System.out.println("Data from Excel is " + data);
            }
*/

/*

        //doc tat ca
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < rowcount; i++)
            for (int j = 0;j<6 ;j++){
                String data = formatter.formatCellValue(Sheet1.getRow(i).getCell(j));
                System.out.println("Data from Excel is " + data);
            }
*/



    //cach 2
    /*
    public static void main(String[] args) throws IOException {

        File src = new File("src/test/resources/loginDataDEMOQA.xlsx");

        FileInputStream fis = new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet1 = wb.getSheetAt(0);
        int rowcount = sheet1.getLastRowNum() + 1; //the number of row
        System.out.println("total rows is: " + rowcount);

        //lay du lieu tung cell
        *//*String data0 = sheet1.getRow(0).getCell(0).getStringCellValue();
        System.out.println("Data from Excel is "+data0);
        String data1 = sheet1.getRow(0).getCell(1).getStringCellValue();
        System.out.println("Data from Excel is "+data1);*//*


        //doc 1 cot
        *//*DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < rowcount; i++) {
            String data0 = formatter.formatCellValue(sheet1.getRow(i).getCell(0));
            System.out.println("Data from Excel is " + data0);
        }*//*




*//*

        //doc tat ca
        DataFormatter formatter = new DataFormatter();
        for (int i = 0; i < rowcount; i++)
            for (int j = 0;j<6 ;j++){
                String data = formatter.formatCellValue(sheet1.getRow(i).getCell(j));
                System.out.println("Data from Excel is " + data);
            }
*//*



    //    wb.close();
    }*/

