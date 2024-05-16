import com.google.api.services.sheets.v4.model.ValueRange
import sheets.SheetsService

fun main() {

    val values = ValueRange().setValues(
        listOf(
            listOf("qwe"),
            listOf("asd", "fgf"),
            listOf("zxcz", "qweqw", "gdfgd")
        )
    )

    SheetsService.update("E4", values)
}