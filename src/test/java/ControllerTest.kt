import org.junit.Assert.*
import org.junit.Test
import java.io.File
/**
 * Created by Necros on 01.03.2016.
 */
class ControllerTest{
    val controller = Controller()
    @Test
    fun whenItWillTakeNotXlsFileParserReturnEmptyObject(){
        val s = controller.parseExcelFile(File("fakeFile").toPath())
     assertEquals(s.price, 0.0,0.0)
    }
}