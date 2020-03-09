package com.smu87.excel.helper

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter


fun main() {
    val inputFile = chooseFile() ?: return
    val inputWorkbook = XSSFWorkbook(inputFile)

    val outputWorkbook = WorkbookProcessor(inputWorkbook).process()

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
