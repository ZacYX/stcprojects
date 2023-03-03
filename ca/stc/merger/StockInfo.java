/**
 * Stock name, increase reason, increase rate and dates in one excel row
 */
package ca.stc.merger;

class StockInfo {
    private String name;        //股票名称
    private String reason;      //涨停原因
    private Double increaseRate; //涨幅
    private Double increaseDates; //连涨天数

    static final String STOCK_NAME_HEADER = "名称";
    static final String STOCK_REASON_HEADER = "涨停原因";
    static final String STOCK_INCREASE_RATE_HEADER = "涨幅";
    static final String STOCK_INCREASE_DATES_HEADER = "连涨";
    static final String CELL_EMPTY_STRING = "--";
    static final Double STOCK_INCREASE_FLAG = 0.09;
    static final String STOCK_REASON_SPLITTER = "+";
    
    String getName() {
        return this.name;
    }
    String getReason() {
        return this.reason;
    }
    Double getIncreaseRate() {
        return this.increaseRate;
    }
    Double getIncreaseDates() {
        return this.increaseDates;
    }
    void setName(String name) {
        this.name = name;
    } 
    void setReason(String reason) {
        this.reason = reason;
    }
    void setIncreaseRate(Double increaseRate) {
        this.increaseRate = increaseRate;
    }
    void setIncreaseDates(Double increaseDates) {
        this.increaseDates = increaseDates;
    }
}
