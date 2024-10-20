package xirr_calculator_stocks.com.example;

import xirr_calculator_stocks.com.example.models.MutualFundDetail;
import xirr_calculator_stocks.com.example.models.StockDetail;
import xirr_calculator_stocks.com.example.utils.SheetService;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.ParseException;
import java.util.*;

public class XIRR_Calculator {
    public static void main(String[] args) throws Exception, IOException, ParseException, GeneralSecurityException {
        Map<String, Sheet> pnlReportSheets = SheetService.getPnLReportSheets();

        // update shyam's stock's XIRR
        List<StockDetail> shyamStocks = SheetService.getStocksFromSheet(pnlReportSheets.get("ShyamPnLSheet"));
        SheetService.updateStockSheet(shyamStocks, "A GRADE");
        SheetService.updateStockSheet(shyamStocks, "B GRADE");
        SheetService.updateStockSheet(shyamStocks, "C GRADE");
        SheetService.updateStockSheet(shyamStocks, "ETF");

        // update mom's stock's XIRR
        List<StockDetail> momStocks = SheetService.getStocksFromSheet(pnlReportSheets.get("MomPnLSheet"));
        SheetService.updateStockSheet(momStocks, "MOM MARKET");

        //update mutual fund
        List<MutualFundDetail> mutualFunds = SheetService.getMutualFundsFromSheet(pnlReportSheets.get("MutualFundPnLSheet"));
        SheetService.updateMutualFundSheet(mutualFunds, "MUTUAL FUND");
    }
}
