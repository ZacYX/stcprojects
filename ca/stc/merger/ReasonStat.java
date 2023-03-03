/**
 * Insert extracted data to the result excel
 */
package ca.stc.merger;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class ReasonStat {
    static final String SHEET_NAME = "首因";
    static final int HEADER_INDEX = 0;
    static final int CATEGORY_INDEX = 0;
    static final int STOCK_LIST_INDEX = 1;

    private FileInputStream fileToBeUpdated; 
    private FileOutputStream fileUpdated;
    private ArrayList<StockInfo> stockInfoList;

    private Workbook reasonStatWorkbook;
    private Sheet reasonStatSheet;
    private Row currentRow;
    private Cell cellWithCategory;
    private Cell cellWithStockList;
    private String category;    //Reason in updater
    private String stockList;
    private Boolean oldCategory;


    ReasonStat(ArrayList<StockInfo> stockInfoList, FileInputStream fileToBeUpdated, FileOutputStream fileUpdated) {
        this.stockInfoList = stockInfoList;
        this.fileToBeUpdated = fileToBeUpdated;
        this.fileUpdated = fileUpdated;
    }
    void prepare() {
        try {
            this.reasonStatWorkbook = new XSSFWorkbook(this.fileToBeUpdated);
        } catch (IOException e) {
            e.printStackTrace();
        }
        this.reasonStatSheet = (Sheet)this.reasonStatWorkbook.getSheet(ReasonStat.SHEET_NAME);
        //Insert a blank column after the first column to the dataSheet, adding 3 to solve outofbounds exception
        this.reasonStatSheet.shiftColumns(1, 
            this.reasonStatSheet.getRow(ReasonStat.HEADER_INDEX).getLastCellNum() + 3, 1);
        Date date = new Date();
        SimpleDateFormat dateFormatForTitle = new SimpleDateFormat("MMdd");
        this.reasonStatSheet.getRow(ReasonStat.HEADER_INDEX).createCell(STOCK_LIST_INDEX).setCellValue(
            dateFormatForTitle.format(date) + " " + this.stockInfoList.size());
    }
    void insert() {
        for (int i = 0; i < this.stockInfoList.size(); i++) {
            this.oldCategory = false;       //Assume it is a new item
            //First row is header, iterate from the second row
            for (int j = 1; j <= this.reasonStatSheet.getLastRowNum(); j++) {
                this.currentRow = this.reasonStatSheet.getRow(j);
                //Get 2 cells
                this.cellWithCategory = this.currentRow.getCell(ReasonStat.CATEGORY_INDEX); 
                this.cellWithStockList = this.currentRow.getCell(ReasonStat.STOCK_LIST_INDEX);
                if (this.cellWithCategory != null) {
                    if (this.cellWithStockList == null) {
                        this.cellWithStockList = this.currentRow.createCell(ReasonStat.STOCK_LIST_INDEX);
                    }
                }
                //Get content of the 2 cells
                this.category = this.cellWithCategory.getStringCellValue();
                this.stockList = this.cellWithStockList.getStringCellValue();
                //Compare reason in arraylist with category in reason statistic excel
                if (stockInfoList.get(i).getReason().equalsIgnoreCase(this.category)) {
                    //Write increase dates that is greater than 1 at the end of each stock name
                    if(stockInfoList.get(i).getIncreaseDates() > 1) {
                        this.stockList += stockInfoList.get(i).getName() 
                            + stockInfoList.get(i).getIncreaseDates().intValue() + "\n";
                    } else {
                        this.stockList += stockInfoList.get(i).getName() + "\n";
                    }
                    this.cellWithStockList.setCellValue(this.stockList);
                    this.oldCategory = true;
                    break;   //Category found, do not need to find the rows left
                }
            }   
            //New category, insert a new row
            if (this.oldCategory == false) {
                Row newRow = this.reasonStatSheet.createRow(this.reasonStatSheet.getLastRowNum() + 1);
                newRow.createCell(ReasonStat.CATEGORY_INDEX).setCellValue(stockInfoList.get(i).getReason());
                if (stockInfoList.get(i).getIncreaseDates() > 1) {
                    newRow.createCell(ReasonStat.STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() 
                        + stockInfoList.get(i).getIncreaseDates().intValue() + "\n"); 
                } else {
                    newRow.createCell(ReasonStat.STOCK_LIST_INDEX).setCellValue(stockInfoList.get(i).getName() + "\n"); 
                }
            }
        }
        
    }
    void destroy() {
        try {
            this.reasonStatWorkbook.write(fileUpdated);
            this.reasonStatWorkbook.close(); 
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    void process() {
        this.prepare();
        this.insert();
        this.destroy();
    }
}