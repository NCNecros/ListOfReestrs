import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.TextArea
import javafx.stage.DirectoryChooser
import net.lingala.zip4j.core.ZipFile
import net.lingala.zip4j.model.FileHeader
import org.apache.poi.hssf.usermodel.HSSFCellStyle
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

    fun click() {
        val directoryChooser = DirectoryChooser()
        directoryChooser.setInitialDirectory(File("d:\\Temp\\26"))
        directoryChooser.title = "Выберите каталог с файлами"
        val file = directoryChooser.showDialog(null)
        if (file != null) {
            var files = Files.list(file.toPath())
            val arr = ArrayList<Schfakt>()
            var countOfFiles: Int = 0
            for (f in files) {
                if (f.toFile().isFile && (f.toFile().name.endsWith("zip") || f.toFile().name.endsWith("ZIP"))) {
                    countOfFiles++
                    val xlsFile = unpackReestr(f.toFile())
                    val smo = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(1)?.value
                    val lpu = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(2)?.value
                    val schetNumber = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(3)?.value?.toInt()
                    if (xlsFile != null) {
                        val inputStream = FileInputStream(xlsFile.toFile())
                        val wb = WorkbookFactory.create(inputStream)
                        val sheetOne = wb.getSheetAt(0)
                        val rowThree = sheetOne.getRow(2)
                        val cell = rowThree.getCell(0)
                        val monts = "(Декабрь|Январь|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь)"
                        val types = "(основной|дополнительный|повторный)"
                        val regexGroups = " к реестру счетов №(\\d{1,5}) от (\\d{2}\\.\\d{2}.\\d{4}) за \\d{4} ($monts) ($types) по .+".toRegex().find(cell.stringCellValue)
                        val dateOfReestr = regexGroups?.groups?.get(2)?.value
                        val month = regexGroups?.groups?.get(3)?.value
                        val typeOfReestr = regexGroups?.groups?.get(5)?.value

                        var typeOfHelp: String = ""
                        val sheetTwo = wb.getSheetAt(1)

                        for (row in arrayOf(15, 17, 22, 24, 19, 26, 30, 31, 35, 36, 37, 38, 39, 40, 41)) {
                            val cellWithPrice = sheetTwo.getRow(row).getCell(9).stringCellValue
                            if (!cellWithPrice.equals("-")) {
                                when (row) {
                                    15, 41 -> {
                                        typeOfHelp = "Стационар"
                                    }
                                    22, 24 -> {
                                        typeOfHelp = "Дневной стационар"
                                    }
                                    else -> {
                                        typeOfHelp = "Поликлиника"
                                    }
                                }
                            }
                        }
                        val price = sheetOne.getRow(20).getCell(13).stringCellValue.replace(" ", "").toDouble()

                        val schet = Schfakt(schetNumber, typeOfReestr, month, dateOfReestr, price, typeOfHelp, smo, lpu)
                        arr.add(schet)

                    }
                    textArea.appendText("Обработан файл:${f.toFile().name}\n")
                }
                textArea.appendText("Обработано\n $countOfFiles файлов")

            }
            try {
                val outStream = FileOutputStream("out.xls")

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
            }

            for (schet in arr) {
                val rowNum = arr.indexOf(schet) + 1
                val row = sheet.createRow(rowNum)
                with(row.createCell(0)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.number?.toString())
                }
                with(row.createCell(1)) {
                    setCellType(Cell.CELL_TYPE_STRING)
                    setCellValue(schet.type)
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
            }
            for (numberOfColumn in 0..7) {
                sheet.autoSizeColumn(numberOfColumn)
            }
            wb.write(outStream)
            outStream.close()
            } catch (e: FileNotFoundException) {
                textArea.appendText(e.message)
                textArea.appendText("\n")
            }

        }
    }

    fun unpackReestr(file: File): Path? {
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


    fun getDataFromXls(file: Path) {

    }
}

data class Schfakt(val number: Int?, val type: String?, val month: String?, val dateOfReestr: String?, val price: Double, val typeOfHelp: String?, val smo: String?, val lpu: String?)
