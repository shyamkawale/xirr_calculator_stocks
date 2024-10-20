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
import java.util.*;

public class GoogleDriveUtils {
    //Application name in firebase
    private static final String APPLICATION_NAME = "XIRR Calculator";
    
    //credentials stored for accessing googledrive apis
    private static final String CREDENTIALS_FILE_PATH = "src/main/resources/client_secret.json";

    // Google Drive file ID for file => https://docs.google.com/spreadsheets/d/1yNT3QWv8PDgDIR6ic6VouTck1T-GpPIUMOHuxLzVPPU/edit?usp=sharing
    private static final String Market_PnL_Report_ID = "1yNT3QWv8PDgDIR6ic6VouTck1T-GpPIUMOHuxLzVPPU";

    // Google Drive file ID for file => https://docs.google.com/spreadsheets/d/1iu1hbr7JjPpO4xwpceFZojvDw6o-EGGDUAOJNNhxZbc/edit?usp=sharing
    private static final String STOCK_MARKET_ID = "1iu1hbr7JjPpO4xwpceFZojvDw6o-EGGDUAOJNNhxZbc";

    private static Drive googleDriveService;
    private static Sheets googleSheetsService;

    static {
		try {
            JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
            List<String> SCOPES = Arrays.asList(
                    DriveScopes.DRIVE, // Full Drive access
                    "https://www.googleapis.com/auth/spreadsheets" // Full Sheets access
            );

			NetHttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

            // Load client secrets from file
            GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
                JSON_FACTORY,
                new FileReader(CREDENTIALS_FILE_PATH));

            // Set up OAuth 2.0 authorization code flow
            GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                httpTransport, JSON_FACTORY, clientSecrets, SCOPES)
                .setAccessType("offline")
                .build();

            // Authorize and get credentials
            Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver())
                .authorize("user");

            googleDriveService = new Drive.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();

            googleSheetsService = new Sheets.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
        } catch (GeneralSecurityException | IOException e) {
            System.out.println("Exception ocurred during Authorization");
            e.printStackTrace();
        }
    }

    // Return Google Drive service using authorized credentials
    public static Drive getDriveService() throws IOException, GeneralSecurityException {
        return googleDriveService;
    }

    // Return Google Sheets service using authorized credentials
    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        return googleSheetsService;
    }

    public static ByteArrayInputStream getFileInputStreamFromDrive() throws Exception {
        // Download file from Google Drive
        OutputStream outputStream = new ByteArrayOutputStream();
        
        googleDriveService
            .files()
            .export(Market_PnL_Report_ID, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            .executeMediaAndDownloadTo(outputStream);

        return new ByteArrayInputStream(((ByteArrayOutputStream) outputStream).toByteArray());
    }

    public static List<List<Object>> getSheetRows(String sheetRange) throws IOException, GeneralSecurityException {
        // Get existing data from the sheet
        ValueRange sheetData = googleSheetsService
            .spreadsheets()
            .values()
            .get(STOCK_MARKET_ID, sheetRange)
            .execute();

        return sheetData.getValues();
    }

    public static void updateSheetRows(List<ValueRange> dataToUpdate) throws IOException{
        BatchUpdateValuesRequest batchUpdateRequest = new BatchUpdateValuesRequest()
            .setValueInputOption("RAW")  // Input raw values (numbers, not strings)
            .setData(dataToUpdate);
    
        googleSheetsService.spreadsheets().values()
            .batchUpdate(STOCK_MARKET_ID, batchUpdateRequest)
            .execute();
    }
}