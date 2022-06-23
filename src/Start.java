import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileWriter;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class Start extends Application {
    @Override
    public void start(Stage stage) throws Exception {
        CheckTrial();
        Parent root = FXMLLoader.load(getClass().getResource("MainPanes/PrintAddress.fxml"));
        Scene scene = new Scene(root);
        stage.setScene(scene);
        stage.show();

    }

    private void CheckTrial() throws Exception {
        File activation = new File(System.getProperty("user.home") + "/AppData/Local/Activation.txt");
        LocalDate expirationdate = LocalDate.of(2023, 1, 20);
        DateTimeFormatter format  = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        if (activation.createNewFile()) {
            FileWriter writer;
            writer = new FileWriter(activation);
            writer.write("Expiration Date: "+expirationdate.format(format));
            writer.close();
        }
//        Iterator var6 = alllines.iterator();
        List<String> activate = Files.readAllLines(activation.toPath());
        for (String line : activate) {
            if (line.contains("Expiration")) {
                expirationdate = LocalDate.parse(line.replace("Expiration Date: ", "").trim(), format);
            }
            System.out.println("Expiration "+expirationdate);
        }
        if(LocalDate.now().isAfter(expirationdate)){
            Stage AuthView = new Stage();
            FXMLLoader Loader = new FXMLLoader();
            Parent root = FXMLLoader.load(getClass().getResource("MainPanes/Expired.fxml"));
            Scene cene = new Scene(root);
            AuthView.setScene(cene);
            AuthView.centerOnScreen();
            AuthView.showAndWait();
            System.exit(0);
        }
        if(LocalDate.now().isAfter(LocalDate.of(2023, 1, 20))){
            Stage AuthView = new Stage();
            FXMLLoader Loader = new FXMLLoader();
            Parent root = FXMLLoader.load(getClass().getResource("MainPanes/Expired.fxml"));
            Scene cene = new Scene(root);
            AuthView.setScene(cene);
            AuthView.centerOnScreen();
            AuthView.showAndWait();
            System.exit(0);
        }



    }


    public static void main(String [] args){
        launch(args);
    }
}
