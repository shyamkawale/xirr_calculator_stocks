package xirr_calculator_stocks.com.example.models;

public class MutualFundDetail {
    public String schemeName;
    public String folioNumber;
    public Double totalUnits;
    public Double buyAmount;
    public Double cmpAmount;
    public Double xirrValue;

    public MutualFundDetail(String schemeName, String folioNumber, Double totalUnits, Double buyAmount, Double cmpAmount, Double xirrValue) {
        this.schemeName = schemeName;
        this.folioNumber = folioNumber;
        this.totalUnits = totalUnits;
        this.buyAmount = buyAmount;
        this.cmpAmount = cmpAmount;
        this.xirrValue = xirrValue;
    }

    public void setSchemeName(String schemeName){
        this.schemeName = schemeName;
    }

    public void setFolioNumber(String folioNumber){
        this.folioNumber = folioNumber;
    }

    public void setTotalUnits(Double totalUnits){
        this.totalUnits = totalUnits;
    }

    public void setBuyAmount(Double buyAmount){
        this.buyAmount = buyAmount;
    }

    public void setCmpAmount(Double cmpAmount){
        this.cmpAmount = cmpAmount;
    }

    public void setXirrValue(Double xirrValue){
        this.xirrValue = xirrValue;
    }

    @Override
    public String toString(){
        return "Scheme Name: " + this.schemeName + " folioNumber: " + this.folioNumber + " units: " + this.totalUnits + " buyAmount: " + this.buyAmount + " cmpAmount: " + this.cmpAmount + " xirr: " + this.xirrValue + "%";
    }

}
