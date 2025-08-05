package projects.reports.reportgenerator;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Main.class.getResource("/projects/reports/reportgenerator/view.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 400, 250);

        try {
            var iconStream = Main.class.getResourceAsStream("/icon.png");
            if (iconStream != null) {
                primaryStage.getIcons().add(new Image(iconStream));
            }
        } catch (Exception e) {
            System.out.println("Could not load application icon: " + e.getMessage());
        }
        primaryStage.setTitle("PPR Excel to PDF Converter");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}

