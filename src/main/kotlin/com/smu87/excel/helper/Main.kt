package com.smu87.excel.helper

import org.apache.logging.log4j.LogManager
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.IOException
import java.text.SimpleDateFormat
import java.util.*
import javax.swing.JFileChooser
import javax.swing.JFrame
import javax.swing.JOptionPane
import javax.swing.UIManager
import javax.swing.filechooser.FileNameExtensionFilter

object Main {

    private val LOG = LogManager.getLogger(Main::class.java)
    private val DATE_FORMAT: SimpleDateFormat = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")

    @JvmStatic
    fun main(args: Array<String>) {
        setSystemLookAndFeel()

        val frame = JFrame()
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        val inputFile = chooseFile(frame)

        try {
            inputFile?.let(this::doWork)
        } catch (e: Throwable) {
            LOG.error(e.message, e)
            inputFile?.let(this::copyFile)
            showErrorDialog(frame)
        } finally {
            frame.dispose()
        }
    }

    private fun copyFile(file: File) {
        val fileName = "${file.nameWithoutExtension} (${DATE_FORMAT.format(Date())}).${file.extension}"
        try {
            val dst = File("problemFiles", fileName)
            file.copyTo(dst)
        } catch (ignore: Throwable) {}
    }

    private fun doWork(inputFile: File) {
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

    private fun chooseFile(frame: JFrame): File? {
        val chooser = JFileChooser()
        chooser.fileSelectionMode = JFileChooser.FILES_ONLY
        chooser.fileFilter = FileNameExtensionFilter("", "xlsx")
        val result = chooser.showOpenDialog(frame)
        return if (result == JFileChooser.APPROVE_OPTION) {
            chooser.selectedFile
        } else {
            null
        }
    }

    private fun showErrorDialog(frame: JFrame) {
        JOptionPane.showMessageDialog(frame, "Что-то пошло не так", "Ошибка", JOptionPane.ERROR_MESSAGE)
    }

    private fun setSystemLookAndFeel() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName())
        } catch (e: Exception) {
            LOG.warn(e.message, e)
        }
    }
}
