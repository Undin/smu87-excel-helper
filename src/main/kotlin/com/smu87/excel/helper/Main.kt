package com.smu87.excel.helper

import org.apache.logging.log4j.LogManager
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.IOException
import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter

object Main {

    private val LOG = LogManager.getLogger(Main::class.java)

    @JvmStatic
    fun main(args: Array<String>) {
        val inputFile = chooseFile() ?: return
        LOG.info("File `${inputFile.absolutePath}` is chosen")
        val inputWorkbook = try {
            XSSFWorkbook(inputFile)
        } catch (e: IOException) {
            LOG.error(e.message, e)
            return
        }

        val outputWorkbook = WorkbookProcessor(inputWorkbook).process()

        try {
            inputWorkbook.close()
            File(inputFile.parent, "${inputFile.nameWithoutExtension}_processed.${inputFile.extension}")
                .outputStream()
                .use { outputWorkbook.write(it) }
        } catch (e: IOException) {
            LOG.error(e.message, e)
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
            LOG.warn(e.message, e)
        }
    }
}
