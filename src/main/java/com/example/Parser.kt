package com.example

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.jsoup.Jsoup
import java.io.File
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.nio.file.Path

class Parser {
    fun parseFileName(f: Path): Triple<String?, String?, Int?> {
        var smo = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(1)?.value
        if (smo.isNullOrBlank()) {
            smo = "(\\d{4})(\\d{5})(\\d{3})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(1)?.value
        }
        var lpu = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(2)?.value
        if (lpu.isNullOrBlank()) {
            lpu = "(\\d{4})(\\d{5})(\\d{3})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(2)?.value

        }
        var schetNumber = "(\\d{4})(\\d{5})(\\d{5})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(3)?.value?.toInt()
        if (schetNumber == null) {
            schetNumber = "(\\d{4})(\\d{5})(\\d{3})\\.(zip|ZIP)".toRegex().find(f.fileName.toString())?.groups?.get(3)?.value?.toInt()
        }
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
                if (cellWithPrice != "-") {
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
    private fun getTypeOfHelp(description: String) : String {
        val listOfHelp = description.split("\n")
        when (listOfHelp[0]) {
            "(I этап) Диспансеризация взрослого населения",
            "медицинские осмотры несовершеннолетних",
            "медицинские осмотры взрослых"-> {
                return "Диспансеризация"
            }


            "Дневной стационар взрослый",
            "Дневной стационар детский",
            "Дневной стационар женской консультации",
            "Стационар дневного пребывания взрослый",
            "Стационар дневного пребывания детский" -> {
                return "Стационарзамещающие"
            }

            "Женская консультация",
            "неотложная помощь взрослому населению",
            "поликлиника (участковая служба) взрослые",
            "поликлиника (участковая служба) дети",
            "Поликлиника взрослая",
            "Поликлиника детская",
            "Центр здоровья взрослый",
            "Стоматология взрослая",
            "Скорая помощь" -> {
                return "Поликлиника"
            }

            "Стационар взрослый",
            "Стационар детский",
            "высокотехнологичная МП взрослым" -> {
                return "Стационар"
            }

        }
        return "не определено"
    }
    fun parseHTMLFileAlt(htmlFile: String): Schfakt {
        val schet = Schfakt()
        try {
            val doc = Jsoup.parse(File(htmlFile), "utf-8")
            val monts = "(декабрь|январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь)"
            val types = "(основной|дополнительный|повторный)"
            val payMethods = "(ОМС|счетам участковой службы|ВМП \\(выс\\.техн\\.МП\\)|доп\\.дисп\\.\\(осмотрам\\) несов\\., сирот|доп\\.дисп\\.\\(осмотрам\\) взр\\. населения|доп\\.дисп\\.\\(осмотрам\\) взр\\. населен)"
            val elements = doc.select("body > div:nth-child(3) > p:nth-child(2)")
            val regexGroups = "к реестру счетов № (\\d{1,5}) от (\\d{2}\\.\\d{2}.\\d{4}) за \\d{4} ($monts) ($types) по ($payMethods)".toRegex().find(elements[0].text())



            schet.dateOfReestr = regexGroups?.groups?.get(2)?.value
            schet.month = regexGroups?.groups?.get(3)?.value
            schet.typeOfReestr = regexGroups?.groups?.get(5)?.value

            val rows = doc.select("table[border=1]")[1].select("thead > tr")


            rows.forEachIndexed { i, element ->
                if (i>0) {
                    if (i == 1) {
                        schet.price = element.select("td:nth-child(10)>p").text().toDouble()
                    }
                    if (!element.select("td:nth-child(1)>p").text().startsWith("---")&&!element.select("td:nth-child(1)>p").text().startsWith("ИТОГО")) {
                        schet.description += element.select("td:nth-child(1)>p").text() + "\n"
                    }
                }

            }



        } catch (e: FileNotFoundException) {
            println("Не удается найти файл: $htmlFile")
            schet.description = "Не удается найти файл"
        } catch (e: IllegalArgumentException) {
            println("Неправильный тип файла: $htmlFile")
            schet.description = "Неправильный тип файла"
        } catch (e: NullPointerException) {
            println("Неправильный тип счет-фактуры")
            schet.description = "Скорая помощь"
            schet.typeOfReestr = "основной"
            schet.typeOfHelp = "Скорая помощь"
        }
        schet.typeOfHelp=getTypeOfHelp(schet.description)
        return schet
    }

    fun parseHTMLFile(xlsFile: String): Schfakt {
        val schet = Schfakt()
        try {
            val doc = Jsoup.parse(File(xlsFile), "utf-8")
            val monts = "(декабрь|январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь)"
            val types = "(основной|дополнительный|повторный)"
            val payMethods = "(ОМС|счетам участковой службы|ВМП \\(выс\\.техн\\.МП\\)|доп\\.дисп\\.\\(осмотрам\\) несов\\., сирот|доп\\.дисп\\.\\(осмотрам\\) взр\\. населения)"
            val elements = doc.select("body > div:nth-child(3) > p:nth-child(2)")
            val regexGroups = "к реестру счетов № (\\d{1,5}) от (\\d{2}\\.\\d{2}.\\d{4}) за \\d{4} ($monts) ($types) по ($payMethods)".toRegex().find(elements[0].text())



            schet.dateOfReestr = regexGroups?.groups?.get(2)?.value
            schet.month = regexGroups?.groups?.get(3)?.value
            schet.typeOfReestr = regexGroups?.groups?.get(5)?.value

            val rows = doc.select("table[border=1]")[2].select("thead > tr")


            rows.forEachIndexed { index, row ->
                run {
                    if (index in arrayOf(2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)) {
                        val cellWithPrice = row.select("td:nth-child(${row.select("td").size - 1})").text()
                        if (cellWithPrice != "0") {
                            if (cellWithPrice != "0.00") {
                                when (index) {
                                    2, 3, 11, 12 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Стационар"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                    13, 14, 17, 18 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Дневной стационар"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                    4, 5, 6, 7, 15, 16, 19, 22, 27 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Поликлиника"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }
                                    8 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Поликлиника"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                    24, 25, 26 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Диспансеризация"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                    28, 29 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "Диспансеризация"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                    30, 31 -> {
                                        val countOfColumns = row.select("td").size
                                        schet.typeOfHelp = "ВМП"
                                        if (countOfColumns == 3) {
                                            schet.description = rows[index - 1].select("td:nth-child(${rows[index - 1].select("td").size - 3})").text()
                                        } else {
                                            schet.description = row.select("td:nth-child(${row.select("td").size - 3})").text()
                                        }
                                    }

                                }
                            }
                        }
                    }

                }
            }
            schet.price = rows[32].select("td:nth-child(3)").text().toDouble()
        } catch (e: FileNotFoundException) {
            println("Не удается найти файл: $xlsFile")
            schet.description = "Не удается найти файл"
        } catch (e: IllegalArgumentException) {
            println("Неправильный тип файла: $xlsFile")
            schet.description = "Неправильный тип файла"
        } catch (e: NullPointerException) {
            println("Неправильный тип счет-фактуры")
            schet.description = "Скорая помощь"
            schet.typeOfReestr = "основной"
            schet.typeOfHelp = "Скорая помощь"
        }
        return schet
    }

    fun parseAmbulanceExcelFile(xlsFile: Path): Schfakt {

        val schet = Schfakt()
        try {
            val inputStream = FileInputStream(xlsFile.toFile())
            val wb = WorkbookFactory.create(inputStream)
            val sheetOne = wb.getSheetAt(0)
            val rowThree = sheetOne.getRow(27)
            val cell = rowThree.getCell(0).stringCellValue + rowThree.getCell(2).stringCellValue
            val monts = "(декабрь|январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь)"
            val types = "(Основной|Дополнительные|Повторные)"
            val regexGroups = "к\\s+реестру\\s+счетов\\s+№(\\d{1,5})\\s+от\\s+(\\d{2}\\.\\d{2}.\\d{4})г\\.\\s+за\\s+\\d{4}\\s+г\\.\\s+($monts)\\s+($types)\\s+по\\s+.+".toRegex().find(cell)

            schet.dateOfReestr = regexGroups?.groups?.get(2)?.value
            schet.month = regexGroups?.groups?.get(3)?.value
            schet.typeOfReestr = regexGroups?.groups?.get(5)?.value
            val sheetTwo = wb.getSheetAt(1)
            schet.typeOfHelp = "Скорая помощь"
            schet.description = "Скорая помощь"
            schet.price = 0.0

        } catch (e: FileNotFoundException) {
            println("Не удается найти файл: $xlsFile")
            schet.description = "Не удается найти файл"
        } catch (e: IllegalArgumentException) {
            println("Неправильный тип файла: $xlsFile")
            schet.description = "Неправильный тип файла"
        } catch (e: NullPointerException) {
            println("Неправильный тип счет-фактуры")
            schet.description = "Скорая помощь"
            schet.typeOfReestr = "основной"
            schet.typeOfHelp = "Скорая помощь"
        }
        return schet
    }
}