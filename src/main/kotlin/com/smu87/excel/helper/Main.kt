package com.smu87.excel.helper

import org.apache.logging.log4j.LogManager
import org.apache.logging.log4j.Logger
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.IOException
import java.text.SimpleDateFormat
import java.util.*
import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter

object Main {

    private val LOG = LogManager.getLogger(Main::class.java)

    @JvmStatic
    fun main(args: Array<String>) {
        val inputFile = chooseFile() ?: return
        Thread.setDefaultUncaughtExceptionHandler(ExceptionHandler(inputFile))
        LOG.info("File `${inputFile.absolutePath}` is chosen")

        val inputWorkbook = XSSFWorkbook(inputFile)
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

    private class ExceptionHandler(private val file: File) : Thread.UncaughtExceptionHandler {
        override fun uncaughtException(t: Thread, e: Throwable) {
            LOG.error(e.message, e)
            val fileName = "${file.nameWithoutExtension} (${DATE_FORMAT.format(Date())}).${file.extension}"
            try {
                val dst = File("problemFiles", fileName)
                file.copyTo(dst)
            } catch (ignore: Exception) {}
        }

        companion object {
            private val LOG: Logger = LogManager.getLogger(ExceptionHandler::class.java)
            private val DATE_FORMAT: SimpleDateFormat = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
        }
    }
}
