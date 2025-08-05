module projects.reports.reportgenerator {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.apache.pdfbox;
    requires org.apache.logging.log4j;
    requires java.desktop;

    opens projects.reports.reportgenerator to javafx.fxml;
    exports projects.reports.reportgenerator;
}