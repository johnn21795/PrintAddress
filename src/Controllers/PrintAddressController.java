package Controllers;

//import Classes.ConnectDB;

import Classes.ModelClass;
import com.jfoenix.controls.JFXButton;
import com.jfoenix.controls.JFXComboBox;
import com.jfoenix.controls.JFXTextField;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Service;
import javafx.concurrent.Task;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.Window;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;
import java.io.*;
import java.net.URL;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class PrintAddressController implements Initializable {
    public JFXButton PrintBut;
    public TableView<ModelClass> Table;
    public TableColumn<ModelClass, String> NameCol;
    public TableColumn<ModelClass, String> QuantityCol;

    public ObservableList<ModelClass> allAddress = FXCollections.observableArrayList();
    public ObservableList<ModelClass> Data = FXCollections.observableArrayList();
    public JFXButton GetBut;
    public JFXTextField MaxRow;
    public TableColumn<ModelClass, String> PhoneCol;
    public TableColumn<ModelClass, String> AddressCol;

    public static File Excel;
    public JFXButton PrintBut2;
    public JFXTextField MinRow;
    public JFXButton PrintManifest;
    public TableColumn<ModelClass, String> DispatchCol;
    public JFXComboBox<String> Dispatch;
    public JFXButton CreateManifest;
    public JFXTextField Sheet;
    public TableColumn<ModelClass, String> NoCol;
    public JFXButton LoadAddress;
    public JFXButton Refresh;
    InputStream in;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        try {

            File destination = new File(System.getProperty("user.home") + "/Desktop/Addresses/");
            destination.mkdirs();


            in = getClass().getClassLoader().getResourceAsStream("./Template.xlsx");
            assert in != null;
            File file = new File(System.getProperty("user.home") + "/Template.xlsx");
            Files.copy(in, file.toPath(), StandardCopyOption.REPLACE_EXISTING);
            in.close();
            MaxRow.setText("100");
            QuantityCol.setCellFactory(TextFieldTableCell.forTableColumn());
            NameCol.setCellFactory(TextFieldTableCell.forTableColumn());
            AddressCol.setCellFactory(TextFieldTableCell.forTableColumn());
            PhoneCol.setCellFactory(TextFieldTableCell.forTableColumn());
            DispatchCol.setCellFactory(TextFieldTableCell.forTableColumn());
            Dispatch.setItems(FXCollections.observableArrayList("","EZE","TITUS","TOCHI","KELVIN"));
            QuantityCol.setEditable(true);
            QuantityCol.setOnEditCancel(event -> {
                Table.setItems(null);
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setHeaderText("");
                alert.setContentText("Press Enter To Save Changes");
                alert.setTitle("Print Address");
                alert.showAndWait();
                Table.setItems(Data);
            });
            NameCol.setOnEditCommit(event -> {
                Data.get(Table.getEditingCell().getRow()).setCol1(event.getNewValue());
                allAddress.clear();
                for(ModelClass x : Data){
                    int labels = Integer.parseInt(x.getCol4());
                    for(int y = 0; y< labels; y++ ){
                        allAddress.add(new ModelClass((y +1),x.getCol1(),x.getCol2(),x.getCol3(),x.getCol4(), Integer.parseInt(x.getCol4()) > 1 ? (y+1)+" OF "+x.getCol4() : "", x.getCol6()));
                    }
                }
            });
            PhoneCol.setOnEditCommit(event -> {
                Data.get(Table.getEditingCell().getRow()).setCol2(event.getNewValue());
                allAddress.clear();
                for(ModelClass x : Data){
                    int labels = Integer.parseInt(x.getCol4());
                    for(int y = 0; y< labels; y++ ){
                        allAddress.add(new ModelClass((y +1),x.getCol1(),x.getCol2(),x.getCol3(),x.getCol4(), Integer.parseInt(x.getCol4()) > 1 ? (y+1)+" OF "+x.getCol4() : "", x.getCol6()));
                    }
                }
            });
            AddressCol.setOnEditCommit(event -> {
                Data.get(Table.getEditingCell().getRow()).setCol3(event.getNewValue());
                allAddress.clear();
                for(ModelClass x : Data){
                    int labels = Integer.parseInt(x.getCol4());
                    for(int y = 0; y< labels; y++ ){
                        allAddress.add(new ModelClass((y +1),x.getCol1(),x.getCol2(),x.getCol3(),x.getCol4(), Integer.parseInt(x.getCol4()) > 1 ? (y+1)+" OF "+x.getCol4() : "", x.getCol6()));
                    }
                }
            });

            QuantityCol.setOnEditCommit(event -> {
                try{
                    int newvalue = Integer.parseInt(event.getNewValue());
                    if(newvalue > 4){
                        Table.setItems(null);
                        Alert alert = new Alert(Alert.AlertType.ERROR);
                        alert.setHeaderText("");
                        alert.setContentText("Number Cannot Exceed 4");
                        alert.setTitle("Print Address");
                        alert.showAndWait();
                        Table.setItems(Data);
                        return;
                    }
                }catch (Exception e){
                    Table.setItems(null);
                    Alert alert = new Alert(Alert.AlertType.ERROR);
                    alert.setHeaderText("");
                    alert.setContentText("Only Number Allowed");
                    alert.setTitle("Print Address");
                    alert.showAndWait();
                    Table.setItems(Data);
                    return;
                }

                Data.get(Table.getEditingCell().getRow()).setCol4(event.getNewValue());
                allAddress.clear();
                for(ModelClass x : Data){
                    int labels = Integer.parseInt(x.getCol4());
                    for(int y = 0; y< labels; y++ ){
                        allAddress.add(new ModelClass((y +1),x.getCol1(),x.getCol2(),x.getCol3(),x.getCol4(), Integer.parseInt(x.getCol4()) > 1 ? (y+1)+" OF "+x.getCol4() : "", x.getCol6()));
                    }
                }

                Table.setItems(null);
                Table.setItems(Data);
            });

            DispatchCol.setOnEditCommit(event -> {
                Data.get(Table.getEditingCell().getRow()).setCol6(event.getNewValue());
                Table.setItems(null);
                Table.setItems(Data);
            });

            MaxRow.setOnAction(event -> {
                try {
                    LoadTable(Excel.getAbsolutePath());
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            });

        } catch (Exception ignored) {

        }
    }

    public void LoadAddress(ActionEvent event) throws Exception{
        if(event.getSource().equals(LoadAddress)){
            File mainFile = new File(System.getProperty("user.home") + "/OneDrive Admin/OneDrive/BVFM/BVFM.NG_4443336_1321418.xlsx");
            File workingFile = new File(System.getProperty("user.home") + "/Desktop/BVFM/BVFM.NG_4443336_1321418.xlsx");
            File workingFile2 = new File(System.getProperty("user.home") + "/Desktop/BVFM/BVFM.NG_4443336_1321417.xlsx");

            Service<Void> Loadservice = new Service<Void>() {
                @Override
                protected Task<Void> createTask() {
                    return new Task<Void>() {
                        @Override
                        protected Void call() throws Exception {
                            Files.copy(mainFile.toPath(), workingFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                            Excel = workingFile;
                            return null;
                        }
                    };
                }
            };


            Loadservice.setOnSucceeded(event1 -> Loadservice.cancel());
            Loadservice.setOnFailed(event1 -> {Loadservice.cancel(); System.out.println("Failed");});
            Loadservice.restart();

            try {


                LoadTable(workingFile.getAbsolutePath());
            } catch (IOException e) {
                try {
                    Files.copy(mainFile.toPath(), workingFile2.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    LoadTable(workingFile2.getAbsolutePath());
                } catch (Exception ex) {
                    //TODO Exit Excel
                    Runtime runtime = Runtime.getRuntime();
                }
               e.printStackTrace();
            }
        } else {
            Parent root = FXMLLoader.load(getClass().getResource("../MainPanes/PrintAddress.fxml"));
            Scene scene = new Scene(root);
            Stage stage = (Stage) LoadAddress.getScene().getWindow();
            stage.setScene(scene);
            stage.show();
        }

    }

    public void LoadTable(String excel) throws Exception{

            Service<Void> Loadservice = new Service<Void>() {
                @Override
                protected Task<Void> createTask() {
                    return new Task<Void>() {
                        @Override
                        protected Void call() throws Exception {

                        try {
                            int minRow = Integer.parseInt(MinRow.getText());
                            int maxRow = Integer.parseInt(MaxRow.getText());
                            XSSFWorkbook workbook = new XSSFWorkbook(excel);
                            XSSFSheet sheet = workbook.getSheetAt(Integer.parseInt(Sheet.getText()));
                            Row row;
                            Cell cell;
                            Data = FXCollections.observableArrayList();
                            int count = 1;
                            for(int x = minRow; x < maxRow; x++){
                                try {
                                    String Name ="";
                                    String Phone="";
                                    String Address="";
                                    boolean isPhone = false;
                                    row = sheet.getRow(x);
                                    cell = row.getCell(11);
                                    try{
                                        Name = cell.getStringCellValue();
                                    }catch (Exception ignored1){
                                        try{
                                            Name = cell.getCellFormula();
                                        }catch (Exception ignored){
                                        }
                                    }
                                    cell = row.getCell(12);
                                    cell.setCellType(1);
                                    try{
                                        Phone = 0+""+cell.getStringCellValue();
                                    }catch (Exception ignored1){
                                        try{
                                            Phone = 0+""+cell.getCellFormula();
                                        }catch (Exception ignored){

                                        }
                                    }
                                    try{
                                        int phone = Integer.parseInt(Phone);
                                    }catch (Exception f){
                                        isPhone = true;
                                    }
                                    if(Phone.length() < 10){
                                        isPhone = false;
                                    }
                                    cell = row.getCell(13);
                                    try{
                                        Address = cell.getStringCellValue();
                                    }catch (Exception ignored1){
                                        try{
                                            Address = cell.getCellFormula();
                                        }catch (Exception ignored){

                                        }
                                    }

                                    ModelClass data = new ModelClass( count, Name,Phone,Address,"0","","");
                                    if(isPhone){
                                        Data.add(data);
                                        count++;
                                    }
                                    workbook = null;
                                }catch (Exception ignored){                }
                            }

                            allAddress.clear();
                            for(ModelClass x : Data){
                                int labels = Integer.parseInt(x.getCol4());
                                for(int y = 0; y< labels; y++ ){
                                    allAddress.add(new ModelClass((y+1),x.getCol1(),x.getCol2(),x.getCol3(),x.getCol4(), Integer.parseInt(x.getCol4()) > 1 ? (y+1)+" OF "+x.getCol4() : "",x.getCol6()));
                                }
                            }
                            Table.setItems(null);
                            NoCol.setCellValueFactory(new PropertyValueFactory<>("num"));
                            NameCol.setCellValueFactory(new PropertyValueFactory<>("col1"));
                            PhoneCol.setCellValueFactory(new PropertyValueFactory<>("col2"));
                            AddressCol.setCellValueFactory(new PropertyValueFactory<>("col3"));
                            QuantityCol.setCellValueFactory(new PropertyValueFactory<>("col4"));
                            DispatchCol.setCellValueFactory(new PropertyValueFactory<>("col6"));

                            Table.setItems(Data);


                        } catch (IOException e) {
                            throw new RuntimeException(e);
                        }
                            return null;
                        }
                    };
                }
            };


            Loadservice.setOnSucceeded(event -> Loadservice.cancel());

            Loadservice.restart();


    }

    public void actionEvents(ActionEvent event) throws Exception {
        if(event.getSource().equals(CreateManifest)){
            DispatchCol.setVisible(true);
            Dispatch.setVisible(true);
            PrintManifest.setVisible(true);
            return;
        }

        if(event.getSource().equals(Dispatch)){
            DispatchCol.setVisible(false);
            for(ModelClass x : Data){
                x.setCol6(Dispatch.getSelectionModel().getSelectedItem());
            }
            Table.setItems(null);
            Table.setItems(Data);
            DispatchCol.setVisible(true);
            return;
        }


        if (event.getSource().equals(GetBut)) {
            try {
                Stage stage =  (Stage) MaxRow.getScene().getWindow();
                stage.setTitle("Select Excel File");
                Window ownerWindow = stage.getOwner();
                FileChooser chooser = new FileChooser();
                FileChooser.ExtensionFilter e = new FileChooser.ExtensionFilter("Excel Files Only", "*.xlsx");
                chooser.getExtensionFilters().add(e);
                Excel = chooser.showOpenDialog(ownerWindow);
                System.out.println(Excel.getAbsolutePath());
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("Successful");
                alert.setContentText("Adjust End Point Value to Show Addresses");
                alert.setHeaderText(null);
                alert.setWidth(300);
                alert.showAndWait();
                MaxRow.setDisable(false);
                MaxRow.requestFocus();

            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            return;
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Address");
        Map<String, CellStyle> cellStyles = new HashMap();
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setFontName("Calibri");
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyles.put("main", cellStyle);

        cellStyle = workbook.createCellStyle();
        font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setFontName("Calibri");
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cellStyles.put("normal", cellStyle);

        cellStyle = workbook.createCellStyle();
        font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setFontName("Calibri");
        font.setBold(true);
        cellStyle.setAlignment((short) 2);
        cellStyle.setFont(font);
        cellStyles.put("center", cellStyle);

        if(event.getSource().equals(PrintBut)){

            Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
            alert.setHeaderText("");
            alert.setContentText("Print Address");
            alert.setTitle("Print Address");
            alert.showAndWait();

            ButtonType result = alert.getResult();
            if (result.equals(ButtonType.OK)) {
                try {

                    int index = 0;
                    System.out.println(allAddress.size());
                    for(int y = 1; y < allAddress.size()+1; y++ ){
                        if(y % 2 != 0){
                            index++;
                            Row row = sheet.createRow(2*y+(index-2));
                            Cell cell = row.createCell(1);
                            cell.setCellStyle(cellStyles.get("main"));
                            cell.setCellValue(allAddress.get(y-1).getCol1());

                            //Phone
                            row = sheet.createRow(2*y+(index-1));
                            cell = row.createCell(1);
                            cell.setCellStyle(cellStyles.get("normal"));
                            cell.setCellValue(allAddress.get(y-1).getCol2());

                            //Address
                            row = sheet.createRow(2*y+(index));
                            cell = row.createCell(1);
                            cell.setCellStyle(cellStyles.get("normal"));
                            cell.setCellValue(allAddress.get(y-1).getCol3());


                            row = sheet.createRow(2*y+(index+1));
                            cell = row.createCell(1);
                            cell.setCellStyle(cellStyles.get("center"));
                            cell.setCellValue(allAddress.get(y-1).getCol5());
                        }else {
                            Row row = sheet.getRow(2*(y-1)+(index-2));
                            Cell cell = row.createCell(3);
                            cell.setCellStyle(cellStyles.get("main"));
                            cell.setCellValue(allAddress.get(y-1).getCol1());

                            //Phone
                            row = sheet.getRow(2*(y-1)+(index-1));
                            cell = row.createCell(3);
                            cell.setCellStyle(cellStyles.get("normal"));
                            cell.setCellValue(allAddress.get(y-1).getCol2());

                            //Address
                            row = sheet.getRow(2*(y-1)+(index));
                            cell = row.createCell(3);
                            cell.setCellStyle(cellStyles.get("normal"));
                            cell.setCellValue(allAddress.get(y-1).getCol3());


                            row = sheet.getRow(2*(y-1)+(index+1));
                            cell = row.createCell(3);
                            cell.setCellStyle(cellStyles.get("center"));
                            cell.setCellValue(allAddress.get(y-1).getCol5());

                        }
                    }

                    sheet.setMargin((short) 0, 0.3);
                    sheet.setMargin((short) 1, 0.3);
                    sheet.setMargin((short) 2, 0.3);
                    sheet.setMargin((short) 3, 0.3);
                    sheet.setMargin((short) 4, 0.3);
                    sheet.setMargin((short) 5, 0.3);

                    sheet.getPrintSetup().setFitWidth((short) 1);

                    sheet.setColumnWidth(1, 12500);
                    sheet.setColumnWidth(0, 800);
                    sheet.setColumnWidth(2, 800);
                    sheet.setColumnWidth(3, 12500);



                    File file = new File(System.getProperty("user.home") + "/Address.xlsx");
                    File destination = new File(System.getProperty("user.home") + "/Desktop/ Address " + LocalDate.now().format(DateTimeFormatter.ofPattern("dd-MMM-yyyy")) + ".xlsx");
                    if (!file.exists()) {
                        file.createNewFile();
                        destination.mkdirs();
                    }

                    FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.home") + "/Address.xlsx");

                    BufferedOutputStream bo = new BufferedOutputStream(outputStream);
                    workbook.write(bo);
                    Files.move(file.toPath(), destination.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    Desktop desktop = Desktop.getDesktop();
                    if (outputStream != null) {
                        try {
                            outputStream.close();
                        } catch (Throwable var30) {
                        }
                    } else {
                        outputStream.close();
                    }


                    //
                    alert = new Alert(Alert.AlertType.CONFIRMATION);
                    alert.setHeaderText("");
                    alert.setContentText("Address Created Print ?");
                    alert.setTitle("Print Address");
                    alert.showAndWait();
                    result = alert.getResult();
                    if (result.equals(ButtonType.OK)) {
                        desktop.open(destination);
                    }


                } catch (Exception var33) {
                    var33.printStackTrace();
                }
            }

        }

        if(event.getSource().equals(PrintBut2)){

            workbook = new XSSFWorkbook(new File(System.getProperty("user.home") + "/Template.xlsx").getAbsolutePath());
//                if(allAddress.size() > 6 ){
                    int pages = allAddress.size() / 6;
                    int extra = allAddress.size() % 6;

                    int pagex = -1;

                    pagex += (pages + extra);

                    for( int x = 8; x > pagex; x--){
                        workbook.removeSheetAt(x);
                    }


//                }
            cellStyle = workbook.createCellStyle();

            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            font.setBold(true);
            cellStyle.setWrapText(true);
            cellStyle.setFont(font);
            cellStyles.put("main", cellStyle);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            cellStyle.setFont(font);
            cellStyle.setWrapText(true);
            cellStyles.put("normal", cellStyle);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            cellStyle.setFont(font);
            cellStyle.setWrapText(true);
            cellStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
            cellStyles.put("address", cellStyle);


            int index = 1;

            System.out.println(allAddress.size());
            for(int y = 1; y < allAddress.size()+1; y++ ){

                int page = (y-1) /6;
                int xx = y + (7*index) - 6 * ((page + 1));
                if(y % 2 != 0){
                    sheet = workbook.getSheetAt(page );

                    Row row = sheet.getRow(xx);
                    Cell cell = row.createCell(1);
                    cell.setCellStyle(cellStyles.get("main"));
                    cell.setCellValue(allAddress.get(y-1).getCol1());

                    //Phone
                    row = sheet.getRow(xx+1);
                    cell = row.createCell(1);
                    cell.setCellStyle(cellStyles.get("normal"));
                    cell.setCellValue(allAddress.get(y-1).getCol2());

                    //Address
                    row = sheet.getRow(xx+2);
                    cell = row.createCell(1);
                    cell.setCellStyle(cellStyles.get("address"));
                    cell.setCellValue(allAddress.get(y-1).getCol3());


                    row = sheet.getRow(xx+3);
                    cell = row.createCell(1);
                    cell.setCellStyle(cellStyles.get("main"));
                    cell.setCellValue(allAddress.get(y-1).getCol5());

                }else {
                    Row row = sheet.getRow(xx-1);
                    Cell cell = row.createCell(4);
                    cell.setCellStyle(cellStyles.get("main"));
                    cell.setCellValue(allAddress.get(y-1).getCol1());

                    //Phone
                    row = sheet.getRow(xx);
                    cell = row.createCell(4);
                    cell.setCellStyle(cellStyles.get("normal"));
                    cell.setCellValue(allAddress.get(y-1).getCol2());

                    //Address
                    row = sheet.getRow(xx+1);
                    cell = row.createCell(4);
                    cell.setCellStyle(cellStyles.get("normal"));
                    cell.setCellValue(allAddress.get(y-1).getCol3());


                    row = sheet.getRow(xx+2);
                    cell = row.createCell(4);
                    cell.setCellStyle(cellStyles.get("main"));
                    cell.setCellValue(allAddress.get(y-1).getCol5());
                    index++;
                    if(index == 4){
                        index = 1;
                    }
                }
            }
            String generatedString = "x"+ new Random().nextInt(99999);
            System.out.println(generatedString);

            File destination = new File(System.getProperty("user.home") + "/Desktop/Addresses/"+generatedString+".xlsx");
            destination.mkdir();
            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.home") + "/PrettyAddress.xlsx");
            BufferedOutputStream bo = new BufferedOutputStream(outputStream);
            workbook.write(bo);
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (Throwable ignored) {
                }
            } else {
                outputStream.close();
            }
            Files.move(new File(System.getProperty("user.home") + "/PrettyAddress.xlsx").toPath(), destination.toPath(), StandardCopyOption.REPLACE_EXISTING);
            Desktop desktop = Desktop.getDesktop();


            desktop.open(destination);

        }

        if(event.getSource().equals(PrintManifest)){
            File file = new File(System.getProperty("user.home") + "/Manifest.xlsx");
            if (!file.exists()) {
                file.createNewFile();
            }
            workbook = new XSSFWorkbook();

            Set<String> dispatchers = new HashSet<>();
            Map<String, Integer> rowCounts = new HashMap<>();


            for( ModelClass x : Data){
                String dispatch = x.getCol6();
                System.out.println(" dispatch "+dispatch);
                dispatchers.add(dispatch);

            }

            System.out.println(" dispatchers "+dispatchers);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            font.setBold(true);
            cellStyle.setFont(font);
            cellStyles.put("main", cellStyle);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            font.setUnderline(FontUnderline.SINGLE);
            cellStyle.setFont(font);
            cellStyles.put("subMain", cellStyle);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            cellStyle.setFont(font);
            cellStyles.put("normal", cellStyle);

            cellStyle = workbook.createCellStyle();
            font = workbook.createFont();
            font.setFontHeightInPoints((short) 14);
            font.setFontName("Calibri");
            cellStyle.setFont(font);
            cellStyles.put("Address", cellStyle);

           for( int x = 0; x < dispatchers.size(); x++){
               workbook.createSheet((String) dispatchers.toArray()[x]);
               sheet = workbook.getSheetAt(x );

               rowCounts.put((String) dispatchers.toArray()[x], 4 );
               //Heading
               Row row = sheet.createRow(0);
               Cell cell = row.createCell(3);
               cell.setCellStyle(cellStyles.get("main"));
               cell.setCellValue("       DELIVERY ORDERS FOR OGADISCOUNT    ");

               //SubHeading
               row = sheet.createRow(1);
               cell = row.createCell(4);
               cell.setCellStyle(cellStyles.get("submain"));
               cell.setCellValue("www.ogadiscount.com");

               //Date
               row = sheet.createRow(2);
               cell = row.createCell(4);
               cell.setCellStyle(cellStyles.get("normal"));
               cell.setCellValue("Date: " + LocalDate.now().format(DateTimeFormatter.ofPattern("dd-MMM-yyyy")));

           }





            int rowCount;

            System.out.println(Data.size());
            for(int y = 1; y < Data.size()+1; y++ ){

                sheet = workbook.getSheet(Data.get(y-1).getCol6());
                rowCount = rowCounts.get(Data.get(y-1).getCol6());
                Row row = sheet.createRow(rowCount);
                Cell cell = row.createCell(0);
                cell.setCellStyle(cellStyles.get("main"));

                cell.setCellValue(  Data.get(y-1).getCol1());
                rowCount++;

                row = sheet.createRow(rowCount);
                cell = row.createCell(0);
                cell.setCellStyle(cellStyles.get("normal"));
                cell.setCellValue(Data.get(y-1).getCol2());
                rowCount++;

                row = sheet.createRow(rowCount);
                cell = row.createCell(0);
                cell.setCellStyle(cellStyles.get("normal"));
                String Address = Data.get(y-1).getCol3();
                int len = Address.length();
                String SubAddress;
                if (len > 70) {
                    SubAddress = Address.substring(0, 70) + "-";
                    cell.setCellValue(SubAddress);
                    rowCount++;

                    row = sheet.createRow(rowCount);
                    cell = row.createCell(0);
                    cell.setCellStyle(cellStyles.get("normal"));
                    if (len > 140) {
                        SubAddress = Address.substring(71, 140) + "-";
                        cell.setCellValue(SubAddress);
                        rowCount++;

                        row = sheet.createRow(rowCount);
                        cell = row.createCell(0);
                        cell.setCellStyle(cellStyles.get("normal"));
                        SubAddress = Address.substring(140) ;
                    } else {
                        SubAddress = Address.substring(70);
                    }
                } else {
                    SubAddress = Address;

                }
                cell.setCellValue(SubAddress);
                rowCount+=2;
                //After Address +1 Row
                row = sheet.createRow(rowCount );
                cell = row.createCell(0);
                cell.setCellValue("");
                rowCounts.put(Data.get(y-1).getCol6(), rowCount);
            }







            File destination = new File(System.getProperty("user.home") + "/Desktop/Manifest/" +LocalDate.now().format(DateTimeFormatter.ofPattern("dd-MMM-yyyy")) + ".xlsx");

            FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.home") + "/Manifest.xlsx");
            BufferedOutputStream bo = new BufferedOutputStream(outputStream);
            workbook.write(bo);
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (Throwable ignored) {
                }
            } else {
                outputStream.close();
            }
            Files.move(new File(System.getProperty("user.home") + "/Manifest.xlsx").toPath(), destination.toPath(), StandardCopyOption.REPLACE_EXISTING);
            Desktop desktop = Desktop.getDesktop();

            desktop.open(destination);







        }


    }
}



