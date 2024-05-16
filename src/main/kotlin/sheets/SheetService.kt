package sheets

import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes.SPREADSHEETS
import com.google.api.services.sheets.v4.model.Spreadsheet
import com.google.api.services.sheets.v4.model.UpdateValuesResponse
import com.google.api.services.sheets.v4.model.ValueRange
import com.google.common.io.Resources
import java.io.File

object SheetsService {

    private const val SPREADSHEET_ID = "1s2SFqaFw8f8g5yBvPvpP6VXzJ_UEUdgYsMLXFstj2wE"
    private val httpTransport by lazy { GoogleNetHttpTransport.newTrustedTransport() }
    private val jsonFactory: GsonFactory by lazy { GsonFactory.getDefaultInstance() }
    private val credential by lazy { credential(httpTransport) }

    private val sheets: Sheets by lazy {
        Sheets.Builder(httpTransport, jsonFactory, credential)
            .setApplicationName("SheetsDSL")
            .build()
    }

    private fun credential(httpTransport: NetHttpTransport): Credential {
        val credentialUrl = Resources.getResource(
            SheetsService::class.java, "/credentials.json"
        )

        val clientSecrets = GoogleClientSecrets.load(
            jsonFactory, credentialUrl.openStream().bufferedReader()
        )

        val flow = GoogleAuthorizationCodeFlow.Builder(
            httpTransport, jsonFactory, clientSecrets, listOf(SPREADSHEETS)
        )
            .setDataStoreFactory(FileDataStoreFactory(File("tokens")))
            .setAccessType("offline")
            .build()

        val receiver = LocalServerReceiver.Builder().setPort(8888).build()

        return AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
    }

    fun create(spreadsheet: Spreadsheet): Spreadsheet =
        sheets.spreadsheets().create(spreadsheet).execute()

    fun getSpreadsheet(spreadsheetId: String): Spreadsheet {
        return sheets.spreadsheets().get(spreadsheetId).execute()
    }

    fun update(range: String, values: ValueRange): UpdateValuesResponse? {
        return sheets.spreadsheets().values()
            .update(SPREADSHEET_ID, range, values)
            .setValueInputOption("RAW")
            .execute()
    }
}