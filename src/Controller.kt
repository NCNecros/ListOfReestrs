import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.TextArea
import javafx.stage.DirectoryChooser
import java.io.File
import java.nio.file.Files
import java.nio.file.Path

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
        directoryChooser.title = "Выберите каталог с файлами"
        val file = directoryChooser.showDialog(null)
        if (file != null) {
            var files = Files.list(file.toPath())
                        .filter {f ->
                        f.fileName.endsWith("zip") && f.fileName.endsWith("ZIP")
                                && (f.fileName.startsWith("1207") || f.fileName.startsWith("4407") || f.fileName.startsWith("1507") || f.fileName.startsWith("9007") || f.fileName.startsWith("1807"))
                    }

                    for (f in files) {
                        textArea.appendText(f.fileName.toString()+"\n")
                    }
        }
    }
}
