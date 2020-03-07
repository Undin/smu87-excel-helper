package com.smu87.excel.helper

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter


// layout:
// number | supplier | name | units | amount | price | cost
fun main() {
    val inputFile = chooseFile() ?: return

    val inputWorkbook = XSSFWorkbook(inputFile)
    val outputWorkbook = XSSFWorkbook()

    processWorkbook(inputWorkbook, outputWorkbook)

    File(inputFile.parent, "${inputFile.nameWithoutExtension}_processed.${inputFile.extension}").outputStream().use {
        outputWorkbook.write(it)
    }
}

private fun chooseFile(): File? {
    setSystemLookAndFeel()

    val frame = JFrame()
    frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
    val chooser = JFileChooser()
    chooser.fileSelectionMode = JFileChooser.FILES_ONLY
    chooser.fileFilter = FileNameExtensionFilter("", "xlsx")
    val result = chooser.showOpenDialog(frame)
    try {
        return if (result == JFileChooser.APPROVE_OPTION) {
            chooser.selectedFile
        } else {
            null
        }
    } finally {
        frame.dispose()
    }
}

private fun setSystemLookAndFeel() {
    try {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName())
    } catch (e: Exception) {
    }
}

private val SECONDARY_SUPPLIER: Regex = """\d*-\d*""".toRegex()
private val UNIT_PATTERN: Regex = """(?<count>\d+) (?<unit>.+)""".toRegex()
private val ROW_COMPARATOR: Comparator<Row> = Comparator.comparing<Row, Boolean> {
    it.getCell(1).value.matches(SECONDARY_SUPPLIER)
}.thenBy {
    it.getCell(2).value
}

private fun processWorkbook(workbook: Workbook, outputWorkbook: Workbook) {
    for (sheet in workbook) {
        val outputSheet = outputWorkbook.createSheet(sheet.sheetName)
        processSheet(sheet, outputSheet)
    }
}

private fun processSheet(sheet: Sheet, outputSheet: Sheet) {
    val titleRow = outputSheet.createRow(0)
    for ((index, columnName) in listOf("№", "Поставщик", "Наименование", "Ед. измерения", "Количество", "Цена", "Сумма").withIndex()) {
        val cell = titleRow.createCell(index, CellType.STRING)
        cell.setCellValue(columnName)
    }

    sheet
        .filter { !it.getCell(4).value.startsWith("-") }
        .sortedWith(ROW_COMPARATOR)
        .forEachIndexed { index, row ->
            val outputRow = outputSheet.createRow(index + 1)
            processRow(row, outputRow)
        }
}

private fun processRow(row: Row, outputRow: Row) {
    var multiplier = 1
    var price: Double? = null
    var amount: Double? = null
    for (i in 0 until 7) {
        when (i) {
            // number
            0 -> {
                val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                outputCell.setCellValue(outputRow.rowNum.toDouble())
            }
            // unit
            3 -> {
                val outputCell = outputRow.createCell(i, CellType.STRING)
                val unit = row[i].value
                val result = UNIT_PATTERN.matchEntire(unit)?.groups
                val unitValue = if (result != null) {
                    multiplier = result["count"]!!.value.toInt()
                    result["unit"]!!.value
                } else {
                    unit
                }
                outputCell.setCellValue(unitValue)
            }
            // amount
            4 -> {
                val value = row[i].value
                amount = value.substringBefore(" x").replace(",", ".").toDouble() * multiplier
                val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                outputCell.setCellValue(amount)
            }
            // price
            5 -> {
                price = row[i].value.toDouble() / multiplier
                val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                outputCell.setCellValue(price)
            }
            // cost
            6 -> {
                val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                outputCell.setCellValue(price!! * amount!!)
            }
            else -> {
                val outputCell = outputRow.createCell(i, CellType.STRING)
                outputCell.setCellValue(row[i].value)
            }
        }
    }
}

private operator fun Row.get(index: Int): Cell? = getCell(index)

private val Cell?.value: String get() {
    return this?.toString().orEmpty().trim()
}
