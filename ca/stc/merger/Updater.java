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
            e.printStackTrace();
        }
        this.updaterSheet = updaterWorkbook.getSheetAt(0); //Updater file has only one sheet
        Row headerRow = this.updaterSheet.getRow(0);       //First row 
        for (Cell cell : headerRow) {
            //There is a space before this string
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_NAME_HEADER)) {
                this.nameColumnIndex = cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_REASON_HEADER)) {
                this.reasonColumnIndex = cell.getColumnIndex();
            }
            //More than one contain this string
            if (cell.getStringCellValue().trim().equals(StockInfo.STOCK_INCREASE_RATE_HEADER)) {
                this.increaseRateColumnIndex= cell.getColumnIndex();
            }
            if (cell.getStringCellValue().trim().contains(StockInfo.STOCK_INCREASE_DATES_HEADER)) {
                this.increaseDatesColumnIndex = cell.getColumnIndex();
            }
            //Stop loop after getting all index
            if (this.nameColumnIndex != -1 && this.reasonColumnIndex != -1 
                    && this.increaseRateColumnIndex != -1 
                    && this.increaseDatesColumnIndex != -1) {
                break;
            }
        }
        System.out.println("Index:  name, " + nameColumnIndex + "    Reason, " + reasonColumnIndex 
            + "    Rate, " + increaseRateColumnIndex + "    dates, " + increaseDatesColumnIndex);
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
            try {
                //Name
                this.stockInfo.setName(this.cellWithName.getStringCellValue());
                //Reason
                this.stockInfo.setReason(this.cellWithReason.getStringCellValue());
                //Increase rate
                if (this.cellWithIncreaseRate.getCellType() == CellType.NUMERIC) {
                    this.stockInfo.setIncreaseRate(this.cellWithIncreaseRate.getNumericCellValue());
                }
                //Increase dates
                if (this.cellWithIncreaseDates.getCellType() == CellType.NUMERIC) {
                    this.stockInfo.setIncreaseDates(this.cellWithIncreaseDates.getNumericCellValue());
                }
            } catch (Exception e) {
                e.printStackTrace();
                continue;
            }
            //Reason: ****+*****+*****+****, get the one before the first "+"
            if (this.stockInfo.getReason().contains(StockInfo.STOCK_REASON_SPLITTER)) {
                this.stockInfo.setReason(this.stockInfo.getReason().substring(0, 
                    this.stockInfo.getReason().indexOf(StockInfo.STOCK_REASON_SPLITTER))) ;
            }
            //Not "--" and increase rate > 0.09 means a valid info, and add it to the list
            if (!this.stockInfo.getReason().equals(StockInfo.CELL_EMPTY_STRING)
                    && this.stockInfo.getIncreaseRate() > StockInfo.STOCK_INCREASE_FLAG
                    && this.stockInfo.getName().length() > 0
                    && this.stockInfo.getIncreaseDates() > 0) {
                this.stockInfoList.add(this.stockInfo);
                this.stockInfo = new StockInfo();
            } else {
                // System.out.println(this.stockInfo.getName() + "    " 
                //                     + this.stockInfo.getReason() + "    " 
                //                     + this.stockInfo.getIncreaseRate() + "    "
                //                     + this.stockInfo.getIncreaseDates());
            }
        } 
        System.out.println("Total items: " + stockInfoList.size());
    }
    void destroy() {
        try {
            this.updaterWorkbook.close();
        } catch (IOException e) {
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