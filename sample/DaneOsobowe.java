package sample;


import javafx.event.ActionEvent;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


public class DaneOsobowe implements HierarchicalController<MainController> {

    public TextField imie;
    public TextField nazwisko;
    public TextField pesel;
    public TextField indeks;
    public TableView<Student> tabelka;
    private MainController parentController;

    public void dodaj(ActionEvent actionEvent) {
        Student st = new Student();
        st.setName(imie.getText());
        st.setSurname(nazwisko.getText());
        st.setPesel(pesel.getText());
        st.setIdx(indeks.getText());
        tabelka.getItems().add(st);
    }

    public void setParentController(MainController parentController) {
        this.parentController = parentController;
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
        tabelka.setEditable(true);
        tabelka.setItems(parentController.getDataContainer().getStudents());
    }

    public void usunZmiany() {
        tabelka.getItems().clear();
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
    }

    public MainController getParentController() {
        return parentController;
    }

    public void initialize() {
        for (TableColumn<Student, ?> studentTableColumn : tabelka.getColumns()) {
            if ("imie".equals(studentTableColumn.getId())) {
                TableColumn<Student, java.lang.String> imieColumn = (TableColumn<Student, java.lang.String>) studentTableColumn;
                imieColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
                imieColumn.setCellFactory(TextFieldTableCell.forTableColumn());
                imieColumn.setOnEditCommit((val)->{
                    val.getTableView().getItems().get(val.getTablePosition().getRow()).setName(val.getNewValue());
                });
            } else if ("nazwisko".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("surname"));
                ((TableColumn<Student, java.lang.String>) studentTableColumn).setCellFactory(TextFieldTableCell.forTableColumn());
            } else if ("pesel".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("pesel"));
                ((TableColumn<Student, java.lang.String>) studentTableColumn).setCellFactory(TextFieldTableCell.forTableColumn());
            } else if ("indeks".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("idx"));
                ((TableColumn<Student, java.lang.String>) studentTableColumn).setCellFactory(TextFieldTableCell.forTableColumn());
            }
        }

    }

  public void synchronizuj() {
        parentController.getDataContainer().setStudents(tabelka.getItems());
    }

    public void dodajJesliEnter(KeyEvent keyEvent) {
        if (keyEvent.getCode() == KeyCode.ENTER) {
            dodaj(new ActionEvent(keyEvent.getSource(), keyEvent.getTarget()));
        }
    }



   public void zapisz(ActionEvent actionEvent) throws IOException {

       synchronizuj();

       HSSFWorkbook workbook = new HSSFWorkbook();
       HSSFSheet sheet = workbook.createSheet("dane");

       //HSSFCellStyle style = workbook.createCellStyle();
       //HSSFFont font = workbook.createFont();
      // font.setBold(true);
       //style.setFont(font);

       HSSFCellStyle stylBold = workbook.createCellStyle();
       HSSFFont bold = workbook.createFont();
       bold.setBold(true);
       stylBold.setFont(bold);

       HSSFRow header = sheet.createRow(0);

       header.createCell(0).setCellValue("Imię");
       header.createCell(1).setCellValue("Nazwisko");
       header.createCell(2).setCellValue("PESEL");
       header.createCell(3).setCellValue("Numer indeksu");
       header.createCell(4).setCellValue("Ocena");
       header.createCell(5).setCellValue("Uzasadnienie");

       for(int i=0; i<6; i++) {
          header.getCell(i).setCellStyle(stylBold);
       }


       int num =1;
       for (Student s : getParentController().getDataContainer().getStudents()) {
           HSSFRow row = sheet.createRow(num);

                      //name
           if (s.getName() != null) {
               row.createCell(0).setCellValue(s.getName());
           } else {
               row.createCell(0).setCellValue("");
           }

           //surname
           if (s.getSurname() != null) {
               row.createCell(1).setCellValue(s.getSurname());
           } else {
               row.createCell(1).setCellValue("");
           }

           //pesel
           if (s.getPesel() != null) {
               row.createCell(2).setCellValue(s.getPesel());
           } else {
               row.createCell(2).setCellValue("");
           }

           //nr indeksu
           if (s.getIdx() != null) {
               row.createCell(3).setCellValue(s.getIdx());
           } else {
               row.createCell(3).setCellValue("");
           }

           //ocena
           HSSFCell cell= row.createCell(4);
           if (s.getGrade() != null) {
               cell.setCellType(cell.CELL_TYPE_NUMERIC);
               cell.setCellValue(s.getGrade().doubleValue());
           //} else {
              // row.createCell(4).setCellValue("");
           }

               //wyjasnienie oceny
               if (s.getGradeDetailed() != null) {
                   row.createCell(5).setCellValue(s.getGradeDetailed());
               } else {
                   row.createCell(5).setCellValue("");
               }


           HSSFCellStyle styl = workbook.createCellStyle();
           styl.setFillPattern(FillPatternType.SOLID_FOREGROUND);
           styl.setFillForegroundColor(IndexedColors.WHITE.getIndex());
           row.setRowStyle(styl);

           HSSFCellStyle kolory = workbook.createCellStyle();
           kolory.setFillPattern(FillPatternType.SOLID_FOREGROUND);
           if (s.getGrade() == null) {
               kolory.setFillForegroundColor(IndexedColors.RED1.getIndex());
           } else if(s.getGrade() < 3.0){
               kolory.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
           } else if(s.getGrade() >= 3.0){
               kolory.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
           }

           row.getCell(4).setCellStyle(kolory);


               num++;
           }

           FileOutputStream fileOut = new FileOutputStream("Dane Osobowe.xlsx");
           workbook.write(fileOut);
           fileOut.close();

       }


//import
 public void wczytaj(ActionEvent actionEvent) throws IOException {
     ArrayList<Student> students = new ArrayList<>();
     try (FileInputStream ois = new FileInputStream("Dane Osobowe.xlsx")) {
         HSSFWorkbook workb = new HSSFWorkbook(ois);
         HSSFSheet sh = workb.getSheet("dane");

         for(int i=1; i<=sh.getLastRowNum(); i++){
             HSSFRow row = sh.getRow(i);
             Student st =new Student();
             st.setName(row.getCell(0).getStringCellValue());
             st.setSurname(row.getCell(1).getStringCellValue());
             st.setPesel(row.getCell(2).getStringCellValue());
             st.setIdx(row.getCell(3).getStringCellValue());
             st.setGradeDetailed(row.getCell(5).getStringCellValue());
            if (row.getCell(4).getNumericCellValue() == 0.0) {
                st.setGrade(null);
            } else{
                st.setGrade(row.getCell(4).getNumericCellValue());
            }

             students.add(st);
         }
         tabelka.getItems().clear();
         tabelka.getItems().addAll(students);
         ois.close();
     } catch (IOException e) {
         e.printStackTrace();
     }
 }
}


