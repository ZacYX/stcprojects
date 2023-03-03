/**
 * Extract stock infomation from the downloaded excel 
 * file from ths.
 */
package ca.stc.merger;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Updater {
    private int nameColumnIndex;
    private int reasonColumnIndex;
    private int increaseRateColumnIndex;
    private int increaseDatesColumnIndex;

    private FileInputStream updaterFile;
    private Workbook updaterWorkbook;
    private Sheet updaterSheet;
    private Row currentRow;
    private Cell cellWithName;
    private Cell cellWithReason;
    private Cell cellWithIncreaseRate;
    private Cell cellWithIncreaseDates;

    private StockInfo stockInfo;
    private ArrayList<StockInfo> stockInfoList;

    Updater(FileInputStream updaterFile) {
        this.updaterFile = updaterFile;
        this.nameColumnIndex = -1;
        this.reasonColumnIndex = -1;
        this.increaseDatesColumnIndex = -1;
        this.increaseRateColumnIndex = -1;
   }

    //prepare workbook, worksheet, collumn index of name, reason, increase rate and dates
    void prepare() {
        this.stockInfo = new StockInfo();
        this.stockInfoList = new ArrayList<StockInfo>();
        try {
            this.updaterWorkbook = new XSSFWorkbook(updaterFile);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        this.updaterSheet = updaterWorkbook.getSheetAt(0); //Updater file has only one sheet
        Row headerRow = this.updaterSheet.getRow(0);       //First row 
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().contains(StockInfo.STOCK_NAME_HEADER)) {
                this.nameColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().contains(StockInfo.STOCK_REASON_HEADER)) {
                this.reasonColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().contains(StockInfo.STOCK_INCREASE_RATE_HEADER)) {
                this.increaseRateColumnIndex= cell.getColumnIndex();
            }
            if (cell.getStringCellValue().contains(StockInfo.STOCK_INCREASE_DATES_HEADER)) {
                this.increaseDatesColumnIndex = cell.getColumnIndex();
            }
            //Stop loop after getting all index
            if (this.nameColumnIndex != -1 && this.reasonColumnIndex != -1 
                    && this.increaseRateColumnIndex != -1 
                    && this.increaseDatesColumnIndex != -1) {
                break;
            }
        }
    }
    //Get stock name, reason, increase rate, increase dates according row index
    void process() {
        for (int i = 1; i < this.updaterSheet.getLastRowNum(); i++) {
            this.currentRow = this.updaterSheet.getRow(i);
            this.cellWithName = this.currentRow.getCell(nameColumnIndex);
            this.cellWithReason = this.currentRow.getCell(reasonColumnIndex);
            this.cellWithIncreaseRate = this.currentRow.getCell(increaseRateColumnIndex);
            this.cellWithIncreaseDates = this.currentRow.getCell(increaseDatesColumnIndex);
            //read cells' content
            if (this.cellWithName != null && this.cellWithReason != null
                    && this.cellWithIncreaseRate != null 
                    && this.cellWithIncreaseDates != null) {
                //Not "--" and increase rate > 0.09 means a valid info, and add it to the list
                if (!this.cellWithReason.getStringCellValue().equals(StockInfo.CELL_EMPTY_STRING)
                        && this.cellWithIncreaseRate.getNumericCellValue() > StockInfo.STOCK_INCREASE_FLAG) {
                    this.stockInfo.setName(this.cellWithName.getStringCellValue());
                    //Reason: ****+*****+*****+****, get the one before the first "+"
                    this.stockInfo.setReason(this.cellWithReason.getStringCellValue().substring(0, 
                        this.cellWithReason.getStringCellValue().indexOf(StockInfo.STOCK_REASON_SPLITTER))) ;
                    this.stockInfo.setIncreaseRate(this.cellWithIncreaseRate.getNumericCellValue());
                    this.stockInfo.setIncreaseDates(this.cellWithIncreaseDates.getNumericCellValue());
                    this.stockInfoList.add(this.stockInfo);
                    this.stockInfo = new StockInfo();
                }
            } else {
                System.out.println("Null pointer in cells of name, reason, increaseRates!");
            }
           
        } 
        System.out.println("Total items: " + stockInfoList.size());
    }
    void destroy() {
        try {
            this.updaterWorkbook.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    ArrayList<StockInfo> getData() {
        this.prepare();
        this.process();
        this.destroy();
        return this.stockInfoList;
    }
}