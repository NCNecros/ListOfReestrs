package com.example
import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.TextArea
import javafx.stage.DirectoryChooser
import net.lingala.zip4j.core.ZipFile
import net.lingala.zip4j.model.FileHeader
import org.apache.commons.io.FileUtils
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFDataFormat
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.util.*

class Controller {
    @FXML
    internal lateinit var button: Button
    @FXML
    internal lateinit var textArea: TextArea
    internal var pathToDir: String = ""
    internal val pathToSettings = Paths.get(System.getProperty("java.io.tmpdir"), "LORSettings")

    internal val tempDirs: ArrayList<Path> = ArrayList()
    internal val parser: Parser = Parser()
    internal val schets: MutableMap<File, String?> = HashMap()


    fun click() {
        val directoryChooser = DirectoryChooser()
        readSettings()
        try {
            directoryChooser.initialDirectory = File(pathToDir)
        }catch (e: Exception){
            directoryChooser.initialDirectory= File("c:\\")
        }
        directoryChooser.title = "Выберите каталог с файлами"
        val dir = directoryChooser.showDialog(null)
        if (dir != null) {
            pathToDir = dir.absolutePath
            saveSettings()
            processDir(dir)
        }
    }

    private fun processDir(file: File) {
        var files = FileUtils.listFiles(file,null,true)
        val a = getListOfFiles(files)
        tempDirs.clear()
        val arr = ArrayList<Schfakt>()
        var countOfFiles: Int = 0
        for (f  in a) {
            countOfFiles++
            val xlsFile = unpackReestr(f)
            val (smo, lpu, schetNumber) = parser.parseFileName(f.toPath())
            val schet = if (xlsFile != null) parser.parseExcelFile(xlsFile) else Schfakt(description = "Счет-фактура отсутствует")
            schet.smo = smo
            schet.lpu = lpu
            schet.schetNumber = schetNumber
            arr.add(schet)
            textArea.appendText("Обработан файл: ${f.name}\n")
            schets.put(f,schet.typeOfReestr)
        }
        textArea.appendText("Обработано $countOfFiles файлов\n")
        saveReport(arr, file)
        removeTempDirs()
        splitToFolders()
    }

    public fun getListOfFiles(files: MutableCollection<File>): List<File> {
        return files.filter {
            (it.name.endsWith("zip") || it.name.endsWith("ZIP")) &&
                    (it.name.startsWith("1207")
                            || it.name.startsWith("1507")
                            || it.name.startsWith("1107")
                            || it.name.startsWith("1807")
                            || it.name.startsWith("9007")
                            || it.name.startsWith("4407")) && !(it.path.contains("По типам"))
        }
    }

    private fun removeTempDirs() {
        for (f in tempDirs) {
            try {
                FileUtils.deleteDirectory(f.toFile())
            } catch (e: Exception) {
                println("Ошибка удаления временного каталога: ${e.message}")
            }
        }
    }
    private fun splitToFolders(){
        val outDir = Paths.get(pathToDir, "По типам")
        if (outDir.toFile().exists()){
            FileUtils.cleanDirectory(outDir.toFile())
        }else{
            outDir.toFile().mkdir()
        }
        for (s in schets){
            when(s.value){
                "основной" -> {
                    val folder = Paths.get(outDir.toString(),  s.key.name.substring(0..3),"Основные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdir()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
                "дополнительный" -> {
                    val folder = Paths.get(outDir.toString(),  s.key.name.substring(0..3),"Дополнительные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdirs()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
                "повторный" -> {
                    val folder = Paths.get(outDir.toString(), s.key.name.substring(0..3),"Повторные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdir()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
            }
        }
    }
    private fun splitToFoldersByType(){
        val outDir = Paths.get(pathToDir, "По типам")
        if (outDir.toFile().exists()){
            FileUtils.cleanDirectory(outDir.toFile())
        }else{
            outDir.toFile().mkdir()
        }
        for (s in schets){
            when(s.value){
                "основной" -> {
                    val folder = Paths.get(outDir.toString(),  s.key.name.substring(0..3),"Основные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdir()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
                "дополнительный" -> {
                    val folder = Paths.get(outDir.toString(),  s.key.name.substring(0..3),"Дополнительные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdirs()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
                "повторный" -> {
                    val folder = Paths.get(outDir.toString(), s.key.name.substring(0..3),"Повторные")
                    if (!folder.toFile().exists()){
                        folder.toFile().mkdir()
                    }
                    FileUtils.copyFile(s.key,Paths.get(folder.toString(),s.key.name).toFile())
                }
            }
        }
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
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.schetNumber?.toString())
                }
                with(row.createCell(1)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.typeOfReestr)
                }
                with(row.createCell(2)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.month)
                }
                with(row.createCell(3)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.dateOfReestr)
                }
                with(row.createCell(4)) {
                    cellType = Cell.CELL_TYPE_NUMERIC
                    setCellValue(schet.price)
                }
                with(row.createCell(5)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.typeOfHelp)
                }
                with(row.createCell(6)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.smo)
                }
                with(row.createCell(7)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.lpu)
                }
                with(row.createCell(8)) {
                    cellType = Cell.CELL_TYPE_STRING
                    setCellValue(schet.description)
                }
            }
            sheet.setAutoFilter(CellRangeAddress(0,arr.size,0,8))
            var lastRowNum = arr.size+2
            val groupedBySMO = arr.groupBy{it.smo}
            for (smo in groupedBySMO.keys){
                val firstRowForMergeSMO = lastRowNum
                val groupedByType = groupedBySMO[smo]!!.groupBy { it.typeOfHelp }
                for (type in groupedByType.keys){
                    val firstRowForMergeByType = lastRowNum
                    val groupedByTypeOfReestr = groupedByType[type]!!.groupBy { it.typeOfReestr }
                    for (typeOfReestr in groupedByTypeOfReestr.keys) {
                        val summ = groupedByTypeOfReestr[typeOfReestr]!!.sumByDouble { it.price }
                        sheet.createRow(lastRowNum).createCell(0, HSSFCell.CELL_TYPE_STRING).setCellValue(smo)
                        sheet.getRow(lastRowNum).createCell(1, HSSFCell.CELL_TYPE_STRING).setCellValue(type)
                        sheet.getRow(lastRowNum).createCell(2, HSSFCell.CELL_TYPE_STRING).setCellValue(typeOfReestr)
                        sheet.getRow(lastRowNum).createCell(3, HSSFCell.CELL_TYPE_NUMERIC).setCellValue(summ)
                        lastRowNum++
                    }
                    sheet.addMergedRegion(CellRangeAddress(firstRowForMergeByType,lastRowNum-1,1,1))
                }
                sheet.addMergedRegion(CellRangeAddress(firstRowForMergeSMO,lastRowNum-1,0,0))
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
        if (!pathToSettings.toFile().exists()) {
            prop.setProperty("pathToDir", "c:\\")
            pathToDir = "c:\\"
            prop.store(FileOutputStream(pathToSettings.toFile()), "comment")
        }
    }

    private fun saveSettings() {
        val prop = Properties()
        prop.setProperty("pathToDir", pathToDir)
        prop.store(FileOutputStream(pathToSettings.toFile()), "commentsss")
    }

    private fun unpackReestr(file: File): Path? {
        val outDir = Files.createTempDirectory("_tmp_${Math.random()}")
        tempDirs.add(outDir)
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
