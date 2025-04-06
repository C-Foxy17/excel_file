package org.example.nurbol_excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelBack {
    private BarChart<String, Number> barChart;
    private ComboBox<String> productComboBox;
    private ObservableList<ProductSale> dataList;
    private TableView<ProductSale> tableView;
    private Stage primaryStage;

    public void start(Stage stage) {
        this.primaryStage = stage;
        stage.setTitle("IS_project_Excel");

        tableView = new TableView<>();
        dataList = FXCollections.observableArrayList();

        TableColumn<ProductSale, Integer> idCol = new TableColumn<>("ID");
        idCol.setCellValueFactory(new PropertyValueFactory<>("id"));

        TableColumn<ProductSale, String> nameCol = new TableColumn<>("Name");
        nameCol.setCellValueFactory(new PropertyValueFactory<>("name"));

        TableColumn<ProductSale, Double> priceCol = new TableColumn<>("Price");
        priceCol.setCellValueFactory(new PropertyValueFactory<>("price"));

        TableColumn<ProductSale, Integer> quantityCol = new TableColumn<>("Quantity");
        quantityCol.setCellValueFactory(new PropertyValueFactory<>("quantity"));

        TableColumn<ProductSale, Double> finalPriceCol = new TableColumn<>("Final Price");
        finalPriceCol.setCellValueFactory(new PropertyValueFactory<>("finalPrice"));

        TableColumn<ProductSale, String> dateCol = new TableColumn<>("Date");
        dateCol.setCellValueFactory(new PropertyValueFactory<>("date"));

        tableView.getColumns().addAll(idCol, nameCol, priceCol, quantityCol, finalPriceCol, dateCol);
        tableView.setPrefHeight(580);

        Button loadButton = new Button("Load .xlsx file");
        loadButton.setOnAction(e -> openFileChooser());
        loadButton.setStyle("-fx-background-color: #2387a6; -fx-padding: 10px 18px; -fx-cursor: pointer; -fx-font-size: 20px;");

        Button showAllButton = new Button("Show all products");
        showAllButton.setOnAction(e -> showAllData());
        showAllButton.setStyle("-fx-background-color: #e8062c; -fx-padding: 10px 18px; -fx-cursor: pointer; -fx-font-size: 20px;");

        productComboBox = new ComboBox<>();
        productComboBox.setPromptText("Choose product");
        productComboBox.setStyle("-fx-background-color: #d5a914; -fx-font-size: 20px; -fx-padding: 8px 18px");
        productComboBox.setOnAction(e -> filterByProduct());

        HBox upPanel = new HBox(10, loadButton, productComboBox, showAllButton);
        upPanel.setPadding(new Insets(10));
        upPanel.setPrefWidth(200);

        CategoryAxis xAxis = new CategoryAxis();
        NumberAxis yAxis = new NumberAxis();
        xAxis.setLabel("Date");
        yAxis.setLabel("Price");

        barChart = new BarChart<>(xAxis, yAxis);
        barChart.setPrefHeight(500);

        BorderPane root = new BorderPane();
        root.setTop(upPanel);
        root.setCenter(tableView);
        root.setBottom(barChart);

        Scene scene = new Scene(root, 1024, 760);
        stage.setScene(scene);
        stage.show();
    }

    private void openFileChooser() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx", "*.xls"));
        File file = fileChooser.showOpenDialog(primaryStage);

        if (file != null) {
            List<ProductSale> loadedData = readExcelFile(file);
            if (!loadedData.isEmpty()) {
                dataList.setAll(loadedData);
                tableView.setItems(dataList);
                Set<String> productNames = loadedData.stream().map(ProductSale::getName).collect(Collectors.toCollection(TreeSet::new));
                productComboBox.getItems().setAll(productNames);
            } else {
                showAlert("File is empty");
            }
        } else {
            showAlert("File is not chosen");
        }
    }

    private void filterByProduct() {
        String selected = productComboBox.getValue();
        if (selected == null) return;
        List<ProductSale> filtered = dataList.stream().filter(p -> p.getName().equalsIgnoreCase(selected)).collect(Collectors.toList());
        tableView.setItems(FXCollections.observableArrayList(filtered));
        updateBarChart(filtered);
    }

    private void updateBarChart(List<ProductSale> filteredData) {
        XYChart.Series<String, Number> series = new XYChart.Series<>();
        series.setName("Product Price");

        for (ProductSale sale : filteredData) {
            series.getData().add(new XYChart.Data<>(sale.getDate(), sale.getFinalPrice()));
        }

        barChart.getData().clear();
        barChart.getData().add(series);
    }

    private void showAllData() {
        tableView.setItems(dataList);
        productComboBox.setValue(null);
        barChart.getData().clear();
    }

    private void showAlert(String message) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle("Error");
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    public static List<ProductSale> readExcelFile(File file) {
        List<ProductSale> result = new ArrayList<>();
        try (FileInputStream inputStream = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                ProductSale sale = new ProductSale();

                Cell idCell = row.getCell(0);
                if (idCell != null && idCell.getCellType() == CellType.NUMERIC) {
                    sale.setId((int) idCell.getNumericCellValue());
                }

                Cell nameCell = row.getCell(1);
                if (nameCell != null && nameCell.getCellType() == CellType.STRING) {
                    sale.setName(nameCell.getStringCellValue());
                }

                Cell priceCell = row.getCell(2);
                if (priceCell != null && priceCell.getCellType() == CellType.NUMERIC) {
                    sale.setPrice(priceCell.getNumericCellValue());
                }

                Cell quantityCell = row.getCell(3);
                if (quantityCell != null && quantityCell.getCellType() == CellType.NUMERIC) {
                    sale.setQuantity((int) quantityCell.getNumericCellValue());
                }

                Cell finalPriceCell = row.getCell(4);
                if (finalPriceCell != null && finalPriceCell.getCellType() == CellType.NUMERIC) {
                    sale.setFinalPrice(finalPriceCell.getNumericCellValue());
                }

                Cell dateCell = row.getCell(5);
                if (dateCell != null) {
                    if (dateCell.getCellType() == CellType.STRING) {
                        sale.setDate(dateCell.getStringCellValue());
                    } else if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                        sale.setDate(dateCell.getDateCellValue().toString());
                    }
                }

                result.add(sale);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }
}