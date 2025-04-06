module org.example.nurbol_excel {
    requires javafx.controls;
    requires javafx.fxml;
    requires static lombok;
    requires org.apache.poi.ooxml;


    opens org.example.nurbol_excel to javafx.fxml;
    exports org.example.nurbol_excel;
}