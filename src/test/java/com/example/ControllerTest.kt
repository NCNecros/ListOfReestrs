package com.example

import org.junit.Test
import org.mockito.Mockito.*;
import org.junit.Assert.*;
import java.io.File

/**
 * Created by Necros on 01.03.2016.
 */
class ControllerTest {
    val controller = Controller()

    @Test
    fun testListOfFilesWithIncorrectIncomingFiles(){
        val list = mutableListOf(File("d:\\1.zip"))

        val result = controller.getListOfFiles(files = list)
        assertEquals(result.size,0)
    }



}