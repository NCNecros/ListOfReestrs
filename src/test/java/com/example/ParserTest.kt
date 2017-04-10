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


    fun whenParserTakeTheCorrectStacionarHTMLFileItWillReturnFilledObject(){
        val s = parser.parseHTMLFile("d:/Temp/2017-01-31/schfakt.html")
        assertEquals(s.dateOfReestr,"31.12.2016")
        assertEquals(s.typeOfReestr,"дополнительный")
        assertEquals(s.typeOfHelp,"Стационар")
        assertEquals(s.price,230697.31,0.0)
        assertEquals(s.description,"Стационар")
    }


    fun whenParserTakeTheCorrectPolicHTMLFileItWillReturnFilledObject(){
        val s = parser.parseHTMLFile("d:/Temp/2017-01-31/schfakt2.html")
        assertEquals(s.dateOfReestr,"31.12.2016")
        assertEquals(s.typeOfReestr,"основной")
        assertEquals(s.typeOfHelp,"Поликлиника")
        assertEquals(s.price,270165.18,0.0)
        assertEquals(s.description,"Поликлиника")
    }
}