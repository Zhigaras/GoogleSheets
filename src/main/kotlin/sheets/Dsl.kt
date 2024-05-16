package sheets

import com.google.api.services.sheets.v4.model.*

class CellDsl(private val value: String? = null) {
    internal fun build() = CellData().apply {
        userEnteredValue = when (value) {
            null, "" -> null
            else -> ExtendedValue().setStringValue(value.toString())
        }
    }
}

class RowDsl(private val height: Int? = null) {

    private val cells = mutableListOf<CellDsl>()

    fun cell(value: String = "", init: CellDsl.() -> Unit = {}) =
        CellDsl(value).apply(init).also { cells += it }

    fun emptyCell(count: Int = 1) {
        if (count < 1) return
        repeat(count) { cell() }
    }

    internal fun build() = RowData().apply {
        setValues(cells.map { it.build() })
    }
}

class SheetDsl {

    private val rows = mutableListOf<RowDsl>()

    fun row(rowCount: Int = 1, init: RowDsl.() -> Unit = {}) {
        repeat(rowCount) {
            rows += RowDsl().apply(init)
        }
    }

    fun build() = Sheet().apply {
        data = listOf(GridData().apply {
            rowData = rows.map { it.build() }
        })
    }
}

class SpreadsheetDsl {

    private val sheetList = mutableListOf<SheetDsl>()

    fun sheet(init: SheetDsl.() -> Unit = {}) {
        sheetList += SheetDsl().apply(init)
    }

    internal fun build() = Spreadsheet().apply {
        sheets = sheetList.map { it.build() }
    }
}

fun spreadsheet(
    init: SpreadsheetDsl.() -> Unit,
): Spreadsheet =
    SheetsService.create(SpreadsheetDsl().apply(init).build())
