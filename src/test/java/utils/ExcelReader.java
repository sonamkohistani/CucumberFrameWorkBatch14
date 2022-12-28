package utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {
    //object one
    static Workbook book;
    //object tow
    static Sheet sheet;
    //this method will open the excel book
    //first method =openExcel
    public static void openExcel(String filePath){
        try {
            FileInputStream fis = new FileInputStream(filePath);

            book = new XSSFWorkbook(fis);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //this method will open the excel work sheet
    //second method get sheet
    public static void getSheet(String sheetName){
        sheet = book.getSheet(sheetName);
    }
    // 3th method getRowCount
    //this method will give the row count
    public static int getRowCount(){
        return sheet.getPhysicalNumberOfRows();
    }
    //4th method getColsCount inside parantacess we use int rowindex
    //this method will give the column count
    public static int getColsCount(int rowIndex){
        return sheet.getRow(rowIndex).getPhysicalNumberOfCells();
    }
   //5th method
    //this method will give the cell data in string format
    public static String getCellData(int rowIndex, int colIndex){
        return sheet.getRow(rowIndex).getCell(colIndex).toString();
    }

//this method will return list of maps having all the data from excel file
    public static List<Map<String, String>> excelListIntoMap
            (String filePath, String sheetName){
        //first method
        openExcel(filePath);
        //second method
        getSheet(sheetName);
        //the data that we want to give from excel file this data save in list of get row
        //creating a list of maps for all the rows
        List<Map<String, String>> listData = new ArrayList<>();
        //one row take care of keys another row take care of value
        //but colunm changing in both the rows
        //loops - outer loop is always take care of rows
        for (int row=1; row<getRowCount(); row++){
            //creating a map for every row
            Map<String, String> map = new LinkedHashMap<>();

            for (int col=0; col<getColsCount(row); col++){
                //this line take care of keys        , //this line take care of value
                map.put(getCellData(0, col), getCellData(row, col));
            }
            //
            listData.add(map);
        }
        return listData;

    }

}