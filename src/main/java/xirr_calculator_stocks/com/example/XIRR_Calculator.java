package xirr_calculator_stocks.com.example;

import xirr_calculator_stocks.com.example.utils.GoogleDriveUtils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.math3.analysis.UnivariateFunction;
import org.apache.commons.math3.analysis.solvers.BrentSolver;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class XIRR_Calculator {
    public static Workbook workbook;
    public static Sheet sheet;

    static {
        try {
            // Google Drive file ID for file => https://docs.google.com/spreadsheets/d/1yNT3QWv8PDgDIR6ic6VouTck1T-GpPIUMOHuxLzVPPU/edit?usp=sharing
            String fileId = "1yNT3QWv8PDgDIR6ic6VouTck1T-GpPIUMOHuxLzVPPU";
            
            // Get file's input Stream from Google Drive and read it
            ByteArrayInputStream file = GoogleDriveUtils.getFileInputStreamFromDrive(fileId);
            workbook = new XSSFWorkbook(file);
            sheet = workbook.getSheetAt(0);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    // Method to read Excel file and get stock data
    public static List<Map<String, Object>> readExcel(String stockName) throws IOException, ParseException {
        List<Map<String, Object>> cashFlows = new ArrayList<>();

        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");

        Date closingDate = new Date();
        double closingAmount = 0;

        // Iterating through rows and filtering by stock name
        for (Row row : sheet) {
            Cell stockCell = row.getCell(0);
            if (stockCell != null && stockCell.getStringCellValue().equals(stockName)) {
                Map<String, Object> cashFlow = new HashMap<>();
                // Buy date and buy value
                Date buyDate = sdf.parse(row.getCell(3).getStringCellValue());
                double buyValue = Double.parseDouble(row.getCell(5).getStringCellValue());

                cashFlow.put("date", buyDate);
                cashFlow.put("cashFlow", -buyValue);  // Buying as negative cash flow
                cashFlows.add(cashFlow);

                // Closing date and closing value
                closingDate = sdf.parse(row.getCell(6).getStringCellValue());
                double closingValue = Double.parseDouble(row.getCell(8).getStringCellValue());

                closingAmount += closingValue;
            }
        }

        Map<String, Object> closingCashFlow = new HashMap<>();
        closingCashFlow.put("date", closingDate);
        closingCashFlow.put("cashFlow", closingAmount);
        cashFlows.add(closingCashFlow);
        return cashFlows;
    }
    
    // XIRR Calculation Method (using Brent's Method)
    public static double calculateXIRR(List<Map<String, Object>> cashFlows, double guess) {
        BrentSolver solver = new BrentSolver(1e-6, 1e-6);  // precision, tolerance

        // Define the XIRR function using UnivariateFunction
        UnivariateFunction xirrFunction = (double rate) -> xirrFunction(cashFlows, rate);

        // Solve for the XIRR using the BrentSolver
        return solver.solve(1000, xirrFunction, -0.9999, 10, guess); // range for rates (-99.99% to very high)
    }

    // Function to calculate the XIRR formula for a given rate (x)
    private static double xirrFunction(List<Map<String, Object>> cashFlows, double rate) {
        double result = 0.0;
        Date startDate = (Date) cashFlows.get(0).get("date");

        for (Map<String, Object> flow : cashFlows) {
            Date date = (Date) flow.get("date");
            double cashFlow = (double) flow.get("cashFlow");

            // Difference in days between the cash flow date and the start date
            double days = (date.getTime() - startDate.getTime()) / (1000.0 * 60 * 60 * 24);
            result += cashFlow / Math.pow(1.0 + rate, days / 365.0);
        }

        return result;
    }

    public static void main(String[] args) throws IOException, ParseException, GeneralSecurityException {
        Set<String> visitedStock = new HashSet<String>();
        double xirrSum = 0;

        List<Map.Entry<String, Double>> xirrList = new ArrayList<>();

        for(Row row : sheet){
            String stockName = row.getCell(0).getStringCellValue();

            if(stockName == "") break;
            if(stockName.contains("Stock")) continue;
            if(visitedStock.contains(stockName)) continue;
            visitedStock.add(stockName);

            // Read the stock data
            List<Map<String, Object>> cashFlows = readExcel(stockName);

            // Initial guess for XIRR calculation
            double xirr = calculateXIRR(cashFlows, 0.1) * 100;
            xirrSum += xirr;

            xirrList.add(new HashMap.SimpleEntry<>(stockName, xirr));
        }
        workbook.close();

        // Sort the list in decreasing order of XIRR
        xirrList.sort((entry1, entry2) -> Double.compare(entry2.getValue(), entry1.getValue()));

        // Print the sorted stocks with their XIRR values
        for (Map.Entry<String, Double> entry : xirrList) {
            System.out.println("XIRR for " + entry.getKey() + ": " + entry.getValue() + "%");
        }

        // Write to Google Sheets
        String spreadsheetId = "1eDj3on2TeZn9PRChoXSW8AZwJBgz3ltVxB14bWEcdg0";  // Replace with actual spreadsheet ID
        String range = "Sheet1!A1";  // Starting range where to write the data
        GoogleDriveUtils.writeXirrResultsToGoogleSheet(xirrList, spreadsheetId, range);

        System.out.println("AVERAGE XIRR: "+ xirrSum/visitedStock.size());
    }
}
