package com.example

import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.nio.file.Path

class Parser {
    fun parseFileName(f: Path): Triple<String?, String?, Int?> {
        val smo = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(1)?.value
        val lpu = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(2)?.value
        val schetNumber = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(3)?.value?.toInt()
        return Triple(smo, lpu, schetNumber)
    }

    fun parseExcelFile(xlsFile: Path): Schfakt {
        val schet = Schfakt()
        try {
            val inputStream = FileInputStream(xlsFile.toFile())
            val wb = WorkbookFactory.create(inputStream)
            val sheetOne = wb.getSheetAt(0)
            val rowThree = sheetOne.getRow(2)
            val cell = rowThree.getCell(0)
            val monts = "(Декабрь|Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь)"
            val types = "(основной|дополнительный|повторный)"
            val regexGroups = " к реестру счетов №(\\d{1,5}) от (\\d{2}\\.\\d{2}.\\d{4}) за \\d{4} ($monts) ($types) по .+".toRegex().find(cell.stringCellValue)

            schet.dateOfReestr = regexGroups?.groups?.get(2)?.value
            schet.month = regexGroups?.groups?.get(3)?.value
            schet.typeOfReestr = regexGroups?.groups?.get(5)?.value
            val sheetTwo = wb.getSheetAt(1)

            for (row in arrayOf(15, 17, 22, 33, 24, 19, 26, 30, 31, 35, 36, 37, 38, 39, 40, 41)) {
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
                        33 -> {
                            schet.typeOfHelp = "ФАП"
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
        } catch (e: FileNotFoundException) {
            println("Не удается найти файл: $xlsFile")
            schet.description = "Не удается найти файл"
        } catch (e: IllegalArgumentException) {
            println("Неправильный тип файла: $xlsFile")
            schet.description = "Неправильный тип файла"
        } catch (e: NullPointerException) {
            println("Неправильный тип счет-фактуры")
            schet.description = "Скорая помощь"
            schet.typeOfHelp = "Скорая помощь"
        }
        return schet
    }
}