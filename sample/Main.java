package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        Parent root = FXMLLoader.load(getClass().getResource("sample.fxml"));
        primaryStage.setTitle("Hello World");
        primaryStage.setScene(new Scene(root,  1056, 555));
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}

// nie zaspisuje po zmianie okna?? ale nie tylko ocen nie zapisuje!!!!!!


/*
Wprowadzaznie danych:

tekst - TextField
liczby - TextField
dane wyboru(płec ited)- Select ComboBox
wartości Boolowskie- CheckBox
listy, tabele- Listy/Tabele

 Statyczne dane-Label

 */
