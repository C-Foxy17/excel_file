package org.example.nurbol_excel;

import javafx.application.Application;
import javafx.stage.Stage;


public class HelloApplication extends Application {
    @Override
    public void start(Stage primaryStage) {
        ExcelBack excelBack = new ExcelBack();
        excelBack.start(primaryStage);
    }

    public static void main(String[] args) {
        launch();
    }
}