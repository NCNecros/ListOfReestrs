import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.TextArea
import javafx.stage.DirectoryChooser
import net.lingala.zip4j.core.ZipFile
import net.lingala.zip4j.model.FileHeader
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.util.*

/**
 * Created by User on 29.02.2016.
 */
class Controller {
    @FXML
    internal lateinit var button: Button
    @FXML
    internal lateinit var textArea: TextArea
    internal var pathToDir: String = ""
    internal val pathToSettings = Paths.get(System.getProperty("java.io.tmpdir"), "LORSettings")
    fun click() {
        val directoryChooser = DirectoryChooser()
        readSettings()
        directoryChooser.setInitialDirectory(File(pathToDir))
        directoryChooser.title = "Выберите каталог с файлами"
        val dir = directoryChooser.showDialog(null)
        if (dir != null) {
            pathToDir = dir.absolutePath
            saveSettings()
            processDir(dir)
        }
    }

    private fun processDir(file: File) {
        var files = Files.list(file.toPath())
        val arr = ArrayList<Schfakt>()
        var countOfFiles: Int = 0
        for (f in files) {
            if (f.toFile().isFile && (f.toFile().name.endsWith("zip") || f.toFile().name.endsWith("ZIP"))) {
                countOfFiles++
                val xlsFile = unpackReestr(f.toFile())
                val (smo, lpu, schetNumber) = parseFileName(f)
                val schet = if (xlsFile != null) parseExcelFile(xlsFile) else Schfakt(description = "Счет-фактура отсутствует")
                schet.smo = smo
                schet.lpu = lpu
                schet.schetNumber = schetNumber
                arr.add(schet)
                textArea.appendText("Обработан файл: ${f.toFile().name}\n")
            }
        }
        textArea.appendText("Обработано $countOfFiles файлов\n")
        saveReport(arr, file)
    }

    private fun saveReport(arr: ArrayList<Schfakt>, file: File) {
        try {
            val outStream = FileOutputStream(Paths.get(file.path, "Итог.xls").toFile())

            val wb = HSSFWorkbook()
            val sheet = wb.createSheet("Итог")
            sheet.printSetup.paperSize = PrintSetup.A4_PAPERSIZE

            with(sheet.createRow(0)) {
                //TODO сделать заголовки жирными
                val font = wb.createFont()
                font.bold = true
                font.fontHeightInPoints = 18

                createCell(0, Cell.CELL_TYPE_STRING).setCellValue("Номер счета")
                createCell(1, Cell.CELL_TYPE_STRING).setCellValue("Тип")
                createCell(2, Cell.CELL_TYPE_STRING).setCellValue("Месяц")
                createCell(3, Cell.CELL_TYPE_STRING).setCellValue("Дата")
                createCell(4, Cell.CELL_TYPE_STRING).setCellValue("Сумма")
                createCell(5, Cell.CELL_TYPE_STRING).setCellValue("Вид помощи")
                createCell(6, Cell.CELL_TYPE_STRING).setCellValue("СМО")
                createCell(7, Cell.CELL_TYPE_STRING).setCellValue("ЛПУ")
                createCell(8, Cell.CELL_TYPE_STRING).setCellValue("Примечание")
            }

            for (schet in arr) {
                val rowNum = arr.indexOf(schet) + 1
                val row = sheet.createRow(rowNum)
                with(row.createCell(0)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.schetNumber?.toString())
                }
                with(row.createCell(1)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.typeOfReestr)
                }
                with(row.createCell(2)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.month)
                }
                with(row.createCell(3)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.dateOfReestr)
                }
                with(row.createCell(4)) {
                    setCellType(Cell.CELL_TYPE_NUMERIC)
                    setCellValue(schet.price)
                }
                with(row.createCell(5)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.typeOfHelp)
                }
                with(row.createCell(6)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.smo)
                }
                with(row.createCell(7)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.lpu)
                }
                with(row.createCell(8)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.description)
                }
            }
            for (numberOfColumn in 0..8) {
                sheet.autoSizeColumn(numberOfColumn)
            }
            wb.write(outStream)
            outStream.close()
        } catch (e: FileNotFoundException) {
            textArea.appendText(e.message)
        }
    }

    private fun parseExcelFile(xlsFile: Path): Schfakt {
        val schet = Schfakt()
        val inputStream = FileInputStream(xlsFile.toFile())
        val wb = WorkbookFactory.create(inputStream)
        val sheetOne = wb.getSheetAt(0)
        val rowThree = sheetOne.getRow(2)
        val cell = rowThree.getCell(0)
        val monts = "(Декабрь|Январь|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь)"
        val types = "(основной|дополнительный|повторный)"
        val regexGroups = " к реестру счетов №(\\d{1,5}) от (\\d{2}\\.\\d{2}.\\d{4}) за \\d{4} ($monts) ($types) по .+".toRegex().find(cell.stringCellValue)

        schet.dateOfReestr = regexGroups?.groups?.get(2)?.value
        schet.month = regexGroups?.groups?.get(3)?.value
        schet.typeOfReestr = regexGroups?.groups?.get(5)?.value
        val sheetTwo = wb.getSheetAt(1)

        for (row in arrayOf(15, 17, 22, 24, 19, 26, 30, 31, 35, 36, 37, 38, 39, 40, 41)) {
            val cellWithPrice = sheetTwo.getRow(row).getCell(9).stringCellValue
            if (!cellWithPrice.equals("-")) {
                when (row) {
                    15, 41 -> {
                        schet.typeOfHelp = "Стационар"
                        schet.description = sheetTwo.getRow(row).getCell(4).stringCellValue
                    }
                    22, 24 -> {
                        schet.typeOfHelp = "Дневной стационар"
                        schet.description = sheetTwo.getRow(row).getCell(4).stringCellValue

                    }
                    else -> {
                        schet.typeOfHelp = "Поликлиника"
                        schet.description = sheetTwo.getRow(row).getCell(4).stringCellValue
                    }
                }
            }
        }
        schet.price = sheetOne.getRow(20).getCell(13).stringCellValue.replace(" ", "").toDouble()
        return schet
    }

    private fun parseFileName(f: Path): Triple<String?, String?, Int?> {
        val smo = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(1)?.value
        val lpu = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(2)?.value
        val schetNumber = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(3)?.value?.toInt()
        return Triple(smo, lpu, schetNumber)
    }

    private fun readSettings() {
        val prop = Properties()
        if (pathToSettings.toFile().exists()) {
            prop.load(FileInputStream(pathToSettings.toFile()))
            pathToDir = prop.getProperty("pathToDir")
        } else {
            createSettings()
        }
    }

    private fun createSettings() {
        val prop = Properties()
        if (!pathToSettings.toFile().exists()){
            prop.setProperty("pathToDir","c:\\")
            pathToDir = "c:\\"
            prop.store(FileOutputStream(pathToSettings.toFile()), "comment")
        }
    }

    private fun saveSettings(){
        val prop = Properties()
        prop.setProperty("pathToDir", pathToDir)
        prop.store(FileOutputStream(pathToSettings.toFile()),"commentsss")
    }

    private fun unpackReestr(file: File): Path? {
        val outDir = Files.createTempDirectory("_tmp_${Math.random()}")
        val zipFile = ZipFile(file)
        var outFile: Path? = null
        for (obj in zipFile.fileHeaders) {
            val header = obj as FileHeader
            if (header.fileName.startsWith("schfakt.xls")) {
                outFile = Paths.get(outDir.toString(), header.fileName)
                zipFile.extractFile(header, outDir.toString())
            }
        }
        return outFile
    }

}

data class Schfakt(var schetNumber: Int? = 0, var typeOfReestr: String? = "", var month: String? = "", var dateOfReestr: String? = "", var price: Double = 0.0, var typeOfHelp: String? = "", var smo: String? = "", var lpu: String? = "", var description: String = "")
