import javafx.application.Application
import javafx.scene.Parent
import javafx.fxml.FXMLLoader.load
import javafx.scene.Scene
import javafx.stage.Stage

/**
 * Created by User on 29.02.2016.
 */
class Main : Application(){
    override fun start(stage: Stage?) {
        val  layout = "layout.fxml"
        stage?.scene = Scene(load<Parent?>(Main.javaClass.getResource(layout)))
        stage?.title = "Отчет по реестрам"
        stage?.show()
    }

    companion object {
        @JvmStatic
        fun main(args: Array<String>){
            launch(Main::class.java)
        }

    }
}