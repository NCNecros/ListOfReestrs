package com.example

import org.junit.Assert.assertEquals
import org.junit.Test
import java.io.File
import java.nio.file.Paths

/**
 * Created by Necros on 02.03.2016.
 */
class ParserTest {
    var parser: Parser = Parser()

    @Test
    fun whenFilenameParserTakeFilenameItWillReturnCorrectData() {
        val (smo, lpu, number) = parser.parseFileName(Paths.get("12070600800228.zip"))
        assertEquals(smo, "1207")
        assertEquals(lpu, "06008")
        assertEquals(number, 228)
    }

    @Test
    fun whenFilenameParserTakeFilenameWithShortNumberItWillReturnCorrectData() {
        val (smo, lpu, number) = parser.parseFileName(Paths.get("120706008228.zip"))
        assertEquals(smo, "1207")
        assertEquals(lpu, "06008")
        assertEquals(number, 228)
    }

    @Test
    fun whenItWillTakeNotFileParserReturnEmptyObject() {
        val s = parser.parseExcelFile(File("fake").toPath())
        assertEquals(s.price, 0.0, 0.0)
        assertEquals(s.schetNumber, 0)
    }

    @Test
    fun whenParserTakeIncorrectXlsFileItWillReturnEmptyObject() {
        val s = parser.parseExcelFile(File(".gitignore").toPath())
        assertEquals(s.price, 0.0, 0.0)
        assertEquals(s.schetNumber, 0)
    }

    @Test
    fun whenPaserTakeCorrectXlsFileItWillReturnFilledObject() {
        val s = parser.parseExcelFile(File(ControllerTest::class.java.getResource("schfakt.xls").file).toPath())
        assertEquals(s.dateOfReestr, "31.01.2016")
        assertEquals(s.typeOfReestr, "основной")
        assertEquals(s.typeOfHelp, "Дневной стационар")
        assertEquals(s.price, 64751.66, 0.0)
    }

    @Test
    fun whenPaserTakeCorrectHTMLFileWithReturnFilledObject() {
        val s = parser.parseHTMLFileAlt(File(ControllerTest::class.java.getResource("schfakt_terapevt.html").file).toString())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "основной")
        assertEquals(s.typeOfHelp, "Поликлиника")
        assertEquals(s.price, 0.0, 0.0)
    }


    @Test
    fun whenPaserTakeCorrectHTMLFileWithAdditionalOMSItWillReturnFilledObject() {
        val s = parser.parseHTMLFileAlt(File(ControllerTest::class.java.getResource("schfakt_dop_oms.html").file).toString())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "дополнительный")
        assertEquals(s.typeOfHelp, "Поликлиника")
        assertEquals(s.price, 3369.9, 0.0)
    }

    @Test
    fun whenPaserTakeCorrectHTMLFileWithMainWillReturnFilledObject() {
        val s = parser.parseHTMLFileAlt(File(ControllerTest::class.java.getResource("schfakt_osn.html").file).toString())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "основной")
        assertEquals(s.typeOfHelp, "Стационар")
        assertEquals(s.price,6537114.33, 0.0)
    }

    @Test
    fun whenPaserTakeCorrectHTMLFileWithPovtOMSWillReturnFilledObject() {
        val s = parser.parseHTMLFileAlt(File(ControllerTest::class.java.getResource("schfakt_povt_oms.html").file).toString())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "повторный")
        assertEquals(s.typeOfHelp, "Стационар")
        assertEquals(s.price, 12226.26, 0.0)
    }

    @Test
    fun whenPaserTakeCorrectHTMLFileWithVMPItWillReturnFilledObject() {
        val s = parser.parseHTMLFileAlt(File(ControllerTest::class.java.getResource("schfakt_vmp.html").file).toString())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "основной")
        assertEquals(s.typeOfHelp, "Стационарзамещающие")
        assertEquals(s.price, 216035.84, 0.0)
    }

    @Test
    fun whenPaserTakeCorrectAmbulanceXlsFileItWillReturnFilledObject() {
        val s = parser.parseAmbulanceExcelFile(File(ControllerTest::class.java.getResource("schfakt_amb.xls").file).toPath())
        assertEquals(s.dateOfReestr, "31.07.2017")
        assertEquals(s.typeOfReestr, "Основной")
        assertEquals(s.typeOfHelp, "Скорая помощь")
        assertEquals(s.price, 2080905.12, 0.0)
    }
}