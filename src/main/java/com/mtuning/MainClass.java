package com.mtuning;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class MainClass extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {
//        Parent root = FXMLLoader.load(getClass().getResource("login.fxml"));
        Parent root = FXMLLoader.load(getClass().getResource("/fxml/login.fxml"));
        primaryStage.setTitle("Logowanie");
        primaryStage.setScene(new Scene(root, 250, 250));
        primaryStage.show();
    }

    public static void main(String[] args) {
        Application.launch(MainClass.class, args);
    }
}