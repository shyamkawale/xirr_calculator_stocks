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

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

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

    // Method to write XIRR results to Google Sheet
    public static void writeXirrResultsToGoogleSheet(Sheets sheetsService, List<Map.Entry<String, Double>> xirrList, String spreadsheetId, String sheetRange) throws IOException, GeneralSecurityException {
        List<ValueRange> dataToUpdate = new ArrayList<>();

        // Prepare data to update the XIRR values for each stock
        ValueRange sheet1Data = sheetsService.spreadsheets().values().get(spreadsheetId, sheetRange).execute();
        List<List<Object>> sheet1Rows = sheet1Data.getValues();
        if (sheet1Rows == null || sheet1Rows.isEmpty()) {
            System.out.println("No data found in the sheet.");
            return;
        }

        int xirrColumnIndex = 15; // XIRR column P

        for (int rowIndex = 2; rowIndex < sheet1Rows.size(); rowIndex++) {
            List<Object> row = sheet1Rows.get(rowIndex);
            String stockName = row.get(1).toString();  // Stock Name is in B column
            for (Map.Entry<String, Double> entry : xirrList) {
                if (entry.getKey().equals(stockName)) {
                    while (row.size() <= xirrColumnIndex) {
                        row.add(null);  // Add empty cells if necessary
                    }
                    sheet1Rows.get(rowIndex).set(xirrColumnIndex, entry.getValue()/100);
                    break;
                }
            }
        }
        dataToUpdate.add(new ValueRange().setRange(sheetRange).setValues(sheet1Rows));

        // Batch update the XIRR values
        if (!dataToUpdate.isEmpty()) {
            BatchUpdateValuesRequest batchUpdateRequest = new BatchUpdateValuesRequest()
                .setValueInputOption("RAW")
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
