package sample;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.ComboBoxTableCell;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;


public class Oceny implements HierarchicalController<MainController> {

    public TextField imie;
    public TextField nazwisko;
    public ComboBox ocena;
    public TextField opisOceny;
    public TableView<Student> tabelka;
    private MainController parentController;
    public ObservableList<Double> ocenki = FXCollections.observableArrayList(1.0,1.5,2.0,2.5,3.0,3.5,4.0,4.5,5.0,5.5,6.0);


    public void dodaj(ActionEvent actionEvent) {
        Student st = new Student();
        st.setName(imie.getText());
        st.setSurname(nazwisko.getText());
        st.setGrade(Double.parseDouble(ocena.getSelectionModel().getSelectedItem().toString()));
        st.setGradeDetailed(opisOceny.getText());
        tabelka.getItems().add(st);
        parentController.getDataContainer().setStudents(tabelka.getItems());
    }

    public void setParentController(MainController parentController) {
        this.parentController = parentController;
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
        tabelka.setEditable(true);
        ocena.getItems().addAll(ocenki);
        tabelka.setItems(parentController.getDataContainer().getStudents());
    }

    public MainController getParentController() {
        return parentController;
    }

    public void initialize() {


        for (TableColumn<Student, ?> studentTableColumn : tabelka.getColumns()) {
            if ("imie".equals(studentTableColumn.getId())) {
                TableColumn<Student, String> imieColumn = (TableColumn<Student, String>) studentTableColumn;
                imieColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
                imieColumn.setCellFactory(TextFieldTableCell.forTableColumn());
                imieColumn.setOnEditCommit((val) -> {
                    val.getTableView().getItems().get(val.getTablePosition().getRow()).setName(val.getNewValue());
                });
            } else if ("nazwisko".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("surname"));
                ((TableColumn<Student, String>) studentTableColumn).setCellFactory(TextFieldTableCell.forTableColumn());
            } else if ("ocena".equals(studentTableColumn.getId())) {
                TableColumn<Student, Double> ocenaColumn = (TableColumn<Student, Double>) studentTableColumn;
                ocenaColumn.setCellValueFactory(new PropertyValueFactory<>("grade"));
                ocenaColumn.setCellFactory(ComboBoxTableCell.forTableColumn(ocenki));
            } else if ("opisOceny".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("gradeDetailed"));
                ((TableColumn<Student, String>) studentTableColumn).setCellFactory(TextFieldTableCell.forTableColumn());
            }
        }

    }


    public void synchronizuj(ActionEvent actionEvent) {
        parentController.getDataContainer().setStudents(tabelka.getItems());
    }



    public void dodajJesliEnter(KeyEvent keyEvent) {
        if (keyEvent.getCode() == KeyCode.ENTER) {
            dodaj(new ActionEvent(keyEvent.getSource(), keyEvent.getTarget()));
        }
    }
}


//////  costam ? costamy : costamx
///jesli costam ==1 to zrób costamy a jesli costam==0 zrób costamx