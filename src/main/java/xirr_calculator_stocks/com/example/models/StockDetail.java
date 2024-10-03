package xirr_calculator_stocks.com.example.models;

public class StockDetail {
    public String stockName;
    public Integer quantity;
    public Double xirrValue;
    public Double avgBuyAmount;
    public Double cmpAmount;

    public StockDetail(String stockName, Integer quantity, Double xirrValue, Double avgBuyAmount, Double cmpAmount) {
        this.stockName = stockName;
        this.quantity = quantity;
        this.xirrValue = xirrValue;
        this.avgBuyAmount = avgBuyAmount;
        this.cmpAmount = cmpAmount;
    }

    public void setQuantity(Integer quantity){
        this.quantity = quantity;
    }

    public void setXirrValue(Double xirrValue){
        this.xirrValue = xirrValue;
    }

    public void setAvgBuyAmount(Double avgBuyAmount){
        this.avgBuyAmount = avgBuyAmount;
    }

    public void setCmpAmount(Double cmpAmount){
        this.cmpAmount = cmpAmount;
    }

    @Override
    public String toString() {
        return "Stock: " + this.stockName + ", Quantity of Stocks: " + this.quantity + ", avgBuyAmount: " + this.avgBuyAmount + ", XIRR: " + this.xirrValue + "%, Current Stock Amount: " + this.cmpAmount;
    }
}
