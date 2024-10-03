package xirr_calculator_stocks.com.example.utils;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import java.security.GeneralSecurityException;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

import xirr_calculator_stocks.com.example.models.StockDetail;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class GoogleDriveUtils {
    private static final String APPLICATION_NAME = "XIRR Calculator";
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();

    private static final List<String> SCOPES = Arrays.asList(
            DriveScopes.DRIVE, // Full Drive access
            "https://www.googleapis.com/auth/spreadsheets" // Full Sheets access
    );
    private static final String CREDENTIALS_FILE_PATH = "src/main/resources/client_secret.json";

    // Load client secrets and get Google Drive service
    public static Drive getDriveService() throws IOException, GeneralSecurityException {
        NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

        // Load client secrets from file
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY,
                new FileReader(CREDENTIALS_FILE_PATH));

        // Set up OAuth 2.0 authorization code flow
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, JSON_FACTORY, clientSecrets, SCOPES)
                .setAccessType("offline")
                .build();

        // Authorize and get credentials
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

        // Return Google Drive service using authorized credentials
        return new Drive.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static ByteArrayInputStream getFileInputStreamFromDrive(String fileId) throws Exception {
        Drive service = getDriveService();

        // Download file from Google Drive
        OutputStream outputStream = new ByteArrayOutputStream();
        
        service.files().export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .executeMediaAndDownloadTo(outputStream);

        return new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
    }

    // Get Google Sheets service
    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

        // Load client secrets from file
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY,
                new FileReader(CREDENTIALS_FILE_PATH));

        // Set up OAuth 2.0 authorization code flow
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, JSON_FACTORY, clientSecrets, SCOPES)
                .setAccessType("offline")
                .build();

        // Authorize and get credentials
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

        // Return Google Sheets service using authorized credentials
        return new Sheets.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static void writeXirrResultsToGoogleSheet(Sheets sheetsService, List<StockDetail> xirrList, String spreadsheetId, String sheetRange) throws IOException, GeneralSecurityException {
        List<ValueRange> dataToUpdate = new ArrayList<>();
    
        // Get existing data from the sheet
        ValueRange sheet1Data = sheetsService.spreadsheets().values().get(spreadsheetId, sheetRange).execute();
        List<List<Object>> sheet1Rows = sheet1Data.getValues();
        if (sheet1Rows == null || sheet1Rows.isEmpty()) {
            System.out.println("No data found in the sheet.");
            return;
        }
    
        // Define the specific columns to update using A1 notation
        String xirrColumnLetter = "P";  // XIRR column P (15th column)
        String avgBuyColumnLetter = "C"; // Average Buy column (3rd column)
        String quantityColumnLetter = "D"; // Quantity column (4th column)
        String cmpAmountColumnLetter = "J"; // CMP column (10th column)
        String sheetName = sheetRange.split("!")[0];
    
        // Iterate through each row in the Google Sheet
        for (int rowIndex = 2; rowIndex < sheet1Rows.size(); rowIndex++) { // Start from row 3 (index 2)
            List<Object> row = sheet1Rows.get(rowIndex);
            String stockName = row.get(1).toString();  // Stock Name is in B column (index 1)
    
            // Find matching stock from the xirrList
            for (StockDetail stock : xirrList) {
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
            BatchUpdateValuesRequest batchUpdateRequest = new BatchUpdateValuesRequest()
                .setValueInputOption("RAW")  // Input raw values (numbers, not strings)
                .setData(dataToUpdate);
    
            sheetsService.spreadsheets().values()
                .batchUpdate(spreadsheetId, batchUpdateRequest)
                .execute();
    
            System.out.println("XIRR values successfully updated.");
        } else {
            System.out.println("No matching stock names found for XIRR update.");
        }
    
        System.out.println("XIRR results written to Google Sheets successfully.");
    }
    
}
