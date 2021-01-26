package com.mtuning;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Variant;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.stage.Stage;

import java.io.IOException;
import java.net.URL;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ResourceBundle;

public class LoginController implements Initializable {

    @FXML
    private Label titleLabel;

    @FXML
    public ComboBox<String> usernameComboBox;

    @FXML
    private PasswordField passwordField;

    @FXML
    private Button loginButton;

    @Override
    public void initialize(URL location, ResourceBundle resourceBundle) {

        ObservableList<String> userList = FXCollections.observableArrayList();

        try {
            DriverManager.registerDriver(new com.microsoft.sqlserver.jdbc.SQLServerDriver());
            Connection connection = DriverManager.getConnection("tajne");

            PreparedStatement preparedStatement = connection.prepareStatement("tajne");
            ResultSet resultSet = preparedStatement.executeQuery();
            while (resultSet.next()) {
                userList.add(resultSet.getString(1));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        usernameComboBox.setItems(userList);
    }

    public void loginButtonOnAction() throws IOException {
        // Logowanie do Subiekt Sfera GT
        ActiveXComponent oGT;
        ActiveXComponent oSubiekt;

        ComThread.InitSTA();
        oGT = new ActiveXComponent("InsERT.GT");
        oGT.setProperty("Produkt", 1);
        oGT.setProperty("Serwer", "tajne");
        oGT.setProperty("Uzytkownik", "sa");
        oGT.setProperty("UzytkownikHaslo", "");
        oGT.setProperty("Baza", "tajne");
//                oGT.setProperty("Operator", username);
//                oGT.setProperty("OperatorHaslo", oDodatki.invoke("Szyfruj", password.toString()));

        oSubiekt = oGT.invokeGetComponent("Uruchom", new Variant(0), new Variant(0));


        try {
            FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/main.fxml"));
            Parent root = (Parent) loader.load();
//            SecondController secondController = loader.getController();
//            secondController.myFunction(textField.getText());

            MainController mainController = loader.getController();
            mainController.setSubiekt(oSubiekt);

            Stage mainstage = new Stage();
            mainstage.setScene(new Scene(root));
            mainstage.show();
            loginButton.getScene().getWindow().hide();

//            Stage mainStage = new Stage();
//            Parent root = FXMLLoader.load(getClass().getResource("/fxml/main.fxml"));
//            mainStage.setTitle("Dodawanie towarów");
//            mainStage.setScene(new Scene(root, 250, 250));
//            mainStage.show();
//            loginButton.getScene().getWindow().hide();

        } catch (IOException e) {
            e.printStackTrace();
        }




//        Stage mainStage = new Stage();
//        Parent root = FXMLLoader.load(getClass().getResource("/fxml/main.fxml"));
//        mainStage.setTitle("Dodawanie towarów");
//        mainStage.setScene(new Scene(root, 250, 250));
//        mainStage.show();
//        loginButton.getScene().getWindow().hide();
    }
}
