package xirr_calculator_stocks.com.example.utils;
import xirr_calculator_stocks.com.example.models.MutualFundDetail;
import xirr_calculator_stocks.com.example.models.StockDetail;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;

public class SheetService {
    public static Map<String, Sheet> getPnLReportSheets() throws Exception{
        Map<String, Sheet> pnlSheets = new HashMap<String, Sheet>();

        // Get file's input Stream from Google Drive and read it
        ByteArrayInputStream pnlReportfile = GoogleDriveUtils.getFileInputStreamFromDrive();
        Workbook workbook = new XSSFWorkbook(pnlReportfile);

        pnlSheets.putIfAbsent("ShyamPnLSheet", workbook.getSheet("ShyamPnLReport"));
        pnlSheets.putIfAbsent("MomPnLSheet", workbook.getSheet("MomPnLReport"));
        pnlSheets.putIfAbsent("MutualFundPnLSheet", workbook.getSheet("MutualFundPnLReport"));

        workbook.close();
        return pnlSheets;
    }

    public static List<StockDetail> getStocksFromSheet(Sheet sheet) throws IOException, ParseException {
        Set<String> visitedStock = new HashSet<String>();
        double xirrSum = 0;
        List<StockDetail> stocks = new ArrayList<>();

        for(Row row : sheet){
            String stockName = row.getCell(0).getStringCellValue();

            if(stockName == "") break;
            if(stockName.contains("Stock") || visitedStock.contains(stockName)) continue;

            StockDetail stock = new StockDetail(stockName, null, null, null, null);
            visitedStock.add(stockName);

            // Read the stock data
            List<Map<String, Object>> cashFlows = getCashFlowForStock(sheet, stock);

            // XIRR calculation
            double xirr = XIRRService.calculateXIRR(cashFlows, 0.1) * 100;
            xirrSum += xirr;
            stock.setXirrValue(xirr);
            
            System.out.println(stock.toString());

            stocks.add(stock);
        }

        System.out.println("AVERAGE XIRR of "+ sheet.getSheetName() +": "+ xirrSum/visitedStock.size());
        return stocks;
    }

    public static void updateStockSheet(List<StockDetail> stocks, String sheetName) throws IOException, GeneralSecurityException{
        List<ValueRange> dataToUpdate = new ArrayList<>();
        String sheetRange = sheetName + "!A1:Z";

        List<List<Object>> sheetRows = GoogleDriveUtils.getSheetRows(sheetRange);
        if (sheetRows == null || sheetRows.isEmpty()) {
            System.out.println("No data found in the sheet.");
            return;
        }
    
        // Define the specific columns to update using A1 notation
        String xirrColumnLetter = "P";  // XIRR column P (15th column)
        String avgBuyColumnLetter = "C"; // Average Buy column (3rd column)
        String quantityColumnLetter = "D"; // Quantity column (4th column)
        String cmpAmountColumnLetter = "J"; // CMP column (10th column)
    
        // Iterate through each row in the Google Sheet
        for (int rowIndex = 2; rowIndex < sheetRows.size(); rowIndex++) { // Start from row 3 (index 2)
            List<Object> row = sheetRows.get(rowIndex);
            String stockName = row.get(1).toString();  // Stock Name is in B column (index 1)
    
            // Find matching stock from the xirrList
            for (StockDetail stock : stocks) {
                if (stock.stockName.equals(stockName)) {
                    List<ValueRange> rowUpdates = new ArrayList<>();
    
                    // Prepare the updates for the specific cells
                    String rowNumber = String.valueOf(rowIndex + 1);  // Row number in A1 notation
    
                    // Update XIRR value in the XIRR column (P)
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + xirrColumnLetter + rowNumber)  // Example: "A GRADE!P3"
                        .setValues(Collections.singletonList(Collections.singletonList(stock.xirrValue / 100))));

                    // Update Average Buy Amount in column C
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + avgBuyColumnLetter + rowNumber)  // Example: "A GRADE!C3"
                        .setValues(Collections.singletonList(Collections.singletonList(stock.avgBuyAmount))));

                    // Update Quantity in column D
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + quantityColumnLetter + rowNumber)  // Example: "A GRADE!D3"
                        .setValues(Collections.singletonList(Collections.singletonList(stock.quantity))));

                    // Update CMP Amount in column J
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + cmpAmountColumnLetter + rowNumber)  // Example: "A GRADE!J3"
                        .setValues(Collections.singletonList(Collections.singletonList(stock.cmpAmount))));

                    // Add updates to the dataToUpdate list
                    dataToUpdate.addAll(rowUpdates);
                    break;  // Exit once the stock is found and updated
                }
            }
        }
    
        // Batch update the specific cells in Google Sheets
        if (!dataToUpdate.isEmpty()) {
            GoogleDriveUtils.updateSheetRows(dataToUpdate);
        } else {
            System.out.println("No matching stock names found for XIRR update.");
        }
        System.out.println("XIRR results updated to Google Sheet" + sheetName + " successfully.");
    }

    // Method to read Excel file and get stock data
    private static List<Map<String, Object>> getCashFlowForStock(Sheet sheet, StockDetail stock) throws IOException, ParseException {
        List<Map<String, Object>> cashFlows = new ArrayList<>();

        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

        Date closingDate = new Date();
        double closingAmount = 0;
        int quantity = 0;
        double totalBuyValue = 0;

        // Iterating through rows and filtering by stock name
        for (Row row : sheet) {
            Cell stockCell = row.getCell(0);
            if (stockCell != null && stockCell.getStringCellValue().equals(stock.stockName)) {
                Map<String, Object> cashFlow = new HashMap<>();
                // FetchData from sheet
                quantity += Integer.parseInt(row.getCell(2).getStringCellValue());
                Date buyDate = sdf.parse(row.getCell(3).getStringCellValue());
                double buyValue = Double.parseDouble(row.getCell(5).toString());

                cashFlow.put("date", buyDate);
                cashFlow.put("cashFlow", -buyValue);  // Buying as negative cash flow
                cashFlows.add(cashFlow);

                // Closing date and closing value
                closingDate = sdf.parse(row.getCell(6).getStringCellValue());
                double closingValue = Double.parseDouble(row.getCell(8).toString());

                closingAmount += closingValue;
                totalBuyValue += buyValue;
            }
        }
        Map<String, Object> closingCashFlow = new HashMap<>();
        closingCashFlow.put("date", closingDate);
        closingCashFlow.put("cashFlow", closingAmount);
        cashFlows.add(closingCashFlow);

        stock.setQuantity(quantity);
        stock.setAvgBuyAmount(totalBuyValue/quantity);
        stock.setCmpAmount(closingAmount/quantity);
        return cashFlows;
    }

    public static List<MutualFundDetail> getMutualFundsFromSheet(Sheet sheet) {
        List<MutualFundDetail> mutualFunds = new ArrayList<>();

        for(Row row : sheet){
            String schemeName = row.getCell(0).getStringCellValue();

            if(schemeName.equals("Scheme Name") || schemeName.equals("")) continue;

            String folioNumber = row.getCell(4).getStringCellValue();
            Double totalUnits = Double.parseDouble(row.getCell(6).toString());
            Double buyAmount = Double.parseDouble(row.getCell(7).toString());
            Double cmpAmount = Double.parseDouble(row.getCell(8).toString());
            Double xirrValue = Double.parseDouble(row.getCell(10).toString().replace("%", "").trim());

            MutualFundDetail mutualFund = new MutualFundDetail(schemeName, folioNumber, totalUnits, buyAmount, cmpAmount, xirrValue);
            
            System.out.println(mutualFund.toString());
            mutualFunds.add(mutualFund);
        }
        return mutualFunds;
    }

    public static void updateMutualFundSheet(List<MutualFundDetail> mutualFunds, String sheetName) throws IOException, GeneralSecurityException {
        List<ValueRange> dataToUpdate = new ArrayList<>();
        String sheetRange = sheetName + "!A1:Z";
        List<List<Object>> sheetRows = GoogleDriveUtils.getSheetRows(sheetRange);
        if (sheetRows == null || sheetRows.isEmpty()) {
            System.out.println("No data found in the sheet.");
            return;
        }

        // Define the specific columns to update using A1 notation
        String unitsColumnLetter = "G"; // Units column (7th column)
        String investedAmountColumnLetter = "H"; // Invested Amount column (8th column)
        String currentAmountColumnLetter = "I"; // CMP column (9th column)
        String xirrColumnLetter = "L"; // Xirr column (12th column)

        // Iterate through each row in the Google Sheet
        for (int rowIndex = 2; rowIndex < sheetRows.size(); rowIndex++) { // Start from row 3 (index 2)
            List<Object> row = sheetRows.get(rowIndex);
            if(row.size() == 0) continue;
            String folioNumber = row.get(2).toString();  // Folio Number is in C column (index 2)
    
            // Find matching stock from the xirrList
            for (MutualFundDetail fund : mutualFunds) {
                if (fund.folioNumber.equals(folioNumber)) {
                    List<ValueRange> rowUpdates = new ArrayList<>();
    
                    // Prepare the updates for the specific cells
                    String rowNumber = String.valueOf(rowIndex + 1);  // Row number in A1 notation

                    // Update Quantity in column D
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + unitsColumnLetter + rowNumber)  // Example: "A GRADE!D3"
                        .setValues(Collections.singletonList(Collections.singletonList(fund.totalUnits))));

                    // Update Average Buy Amount in column C
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + investedAmountColumnLetter + rowNumber)  // Example: "A GRADE!C3"
                        .setValues(Collections.singletonList(Collections.singletonList(fund.buyAmount))));

                    // Update CMP Amount in column J
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + currentAmountColumnLetter + rowNumber)  // Example: "A GRADE!J3"
                        .setValues(Collections.singletonList(Collections.singletonList(fund.cmpAmount))));
    
                    // Update XIRR value in the XIRR column (P)
                    rowUpdates.add(new ValueRange()
                        .setRange(sheetName + "!" + xirrColumnLetter + rowNumber)  // Example: "A GRADE!P3"
                        .setValues(Collections.singletonList(Collections.singletonList(fund.xirrValue / 100))));

                    // Add updates to the dataToUpdate list
                    dataToUpdate.addAll(rowUpdates);
                    break;  // Exit once the stock is found and updated
                }
            }
        }

        // Batch update the specific cells in Google Sheets
        if (!dataToUpdate.isEmpty()) {
            GoogleDriveUtils.updateSheetRows(dataToUpdate);
        } else {
            System.out.println("No matching mutual fund names found for update.");
        }
        System.out.println("Mutual Funds results updated to Google Sheet" + sheetName + " successfully.");
    }
}
