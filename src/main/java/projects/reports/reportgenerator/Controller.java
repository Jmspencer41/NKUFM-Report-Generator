package projects.reports.reportgenerator;

import java.awt.Color;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Controller {
    private static final Logger logger = LogManager.getLogger(Controller.class);

    @FXML
    private TextField rowNumberField;
    @FXML
    private Button selectFileButton;
    @FXML
    private Button convertButton;
    @FXML
    private Label statusLabel;

    private File selectedFile;

    private String projectManager;
    private String projectNumber;
    private String projectName;
    private String projectDescription;
    private String comments;
    private String workPhase;
    private String startDate;
    private String endDate;
    private String architect;
    private String engineer;
    private String contractor;
    private int internalOrder;
    private int costCenter;
    private int fundFY;
    private float percentComplete;
    private float budgetTotal;
    private float budgetUsed;
    private float budgetRemaining;

    @FXML
    private void initialize() throws IOException {
        selectFileButton.setOnAction(event -> selectExcelFile());
        convertButton.setOnAction(event -> convertToPDF());
    }

    private void selectExcelFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            statusLabel.setText("Selected: " + selectedFile.getName());
        } else {
            statusLabel.setText("No file selected!");
        }
    }

    private void convertToPDF() {
        if (selectedFile != null) {
            createPDF();
        }
        else {
            statusLabel.setText("Please select an Excel file!");
        }

    }

    private void createPDF() {
        try (FileInputStream fis = new FileInputStream(selectedFile)) {

            int rowCount;
            Workbook workbook = WorkbookFactory.create(fis);
            //get the number of rows used in the file
            rowCount = assignRowCount(workbook);
            //loop through the columns to get variables and convert to the PDF
            try (PDDocument document = new PDDocument()) {
                for (int i = 1; i < rowCount; i++) {
                    Row row = workbook.getSheetAt(0).getRow(i);
                    assignVariables(row);
                    createPage(document);
                }
                outputFile(document);
            }
            catch (IOException e) {
                statusLabel.setText("Error: Failed to process file.");
                logger.error("Error processing file", e);
            }
        } catch (IOException e) {
            statusLabel.setText("Error: Failed to process file.");
            logger.error("Error processing file", e);
        }

    }
    private int assignRowCount(Workbook workbook) {
        int rowCount = 0;
        for (int i = 0; i <= workbook.getSheetAt(0).getLastRowNum(); i++) {
            Row row = workbook.getSheetAt(0).getRow(i);
            if (row != null) {
                rowCount++;
            }
        }
        return rowCount;
    }

    private void assignVariables(Row row) {
        projectManager = cleanString(getCellString(row, 0, "Unknown"));
        projectNumber = cleanString(getCellString(row, 1, "N/A"));
        projectName = cleanString(getCellString(row, 2, "Untitled"));
        projectDescription = cleanString(getCellString(row, 3, ""));
        comments = cleanString(getCellString(row, 4, ""));
        workPhase = cleanString(getCellString(row, 5, "N/A"));
        percentComplete = getCellFloat(row, 6, 0);
        startDate = cleanString(getCellString(row, 7, "N/A"));
        endDate = cleanString(getCellString(row, 8, "N/A"));
        costCenter = getCellInt(row, 9, 0);
        internalOrder = getCellInt(row, 10, 0);
        fundFY = getCellInt(row, 11, 0);
        budgetTotal = getCellFloat(row, 12, 0.0f);
        budgetUsed = getCellFloat(row, 13, 0.0f);
        budgetRemaining = getCellFloat(row, 14, 0.0f);
        architect = cleanString(getCellString(row, 15, "N/A"));
        engineer = cleanString(getCellString(row, 16, "N/A"));
        contractor = cleanString(getCellString(row, 17, "N/A"));

        percentComplete = percentComplete * 100;
    }

    private String cleanString(String input) {
        return input.replaceAll("[\\n\\r]", " ").trim();
    }

    private String getCellString(Row row, int columnIndex, String defaultValue) {
        Cell cell = row.getCell(columnIndex);
        return cell == null || cell.toString().trim().isEmpty() ? defaultValue : cell.toString().trim();
    }

    private int getCellInt(Row row, int columnIndex, int defaultValue) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null || cell.toString().trim().isEmpty()) {
            return defaultValue;
        }
        try {
            return (int) cell.getNumericCellValue();
        } catch (Exception e) {
            logger.warn("Invalid integer in cell ({}, {}): {}", row.getRowNum(), columnIndex, e.getMessage());
            return defaultValue;
        }
    }

    private float getCellFloat(Row row, int columnIndex, float defaultValue) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null || cell.toString().trim().isEmpty()) {
            return defaultValue;
        }
        try {
            return (float) cell.getNumericCellValue();
        } catch (Exception e) {
            logger.warn("Invalid float in cell ({}, {}): {}", row.getRowNum(), columnIndex, e.getMessage());
            return defaultValue;
        }
    }

    private void createPage(PDDocument document) {
        int docWidth = 792;
        int docHeight = 612;
        PDPage page = new PDPage(new PDRectangle(docWidth, docHeight)); // US Letter landscape
        document.addPage(page);

        try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
            PDType1Font boldFont = new PDType1Font(Standard14Fonts.FontName.HELVETICA_BOLD);
            PDType1Font regularFont = new PDType1Font(Standard14Fonts.FontName.HELVETICA);
            float margin = 30;
            float yStart = 575; // Top of page

            // Header: Project Name and Page Number
            Color headerColor = new Color(Integer.parseInt("FFC72C", 16));
            contentStream.setNonStrokingColor(headerColor);
            float pageWidth = page.getMediaBox().getWidth();
            contentStream.addRect(0, page.getMediaBox().getHeight() - 60, pageWidth, 60);
            contentStream.fill();
            contentStream.setFont(boldFont, 20);
            contentStream.setNonStrokingColor(Color.BLACK);
            contentStream.beginText();
            contentStream.newLineAtOffset(margin - 15, yStart);
            contentStream.showText(String.format("%s: ", projectManager));
            contentStream.showText(projectName);
            contentStream.endText();

            //Project Number Line
            contentStream.setFont(boldFont, 18);
            contentStream.beginText();
            contentStream.newLineAtOffset((int) ((margin * 2) + (docWidth / 2)), 505);
            contentStream.showText(String.format("Project Number: %s", projectNumber));
            contentStream.endText();

            // Description Start Box
            contentStream.setFont(boldFont, 16);
            contentStream.beginText();
            contentStream.newLineAtOffset(margin, 518);
            contentStream.showText("Description: ");
            contentStream.endText();
            wrapText(contentStream, regularFont, margin, 500, 14,docWidth, projectDescription);

            // Comments Start Box
            contentStream.setFont(boldFont, 16);
            contentStream.beginText();
            contentStream.newLineAtOffset(margin, 348);
            contentStream.showText("Status Update: ");
            contentStream.endText();
            wrapText(contentStream, regularFont, margin, 330, 14, docWidth, comments);

            // Info Start Box
            int infoBoxYStart = 485;
            int infoBoxXStart = (int) ((margin * 2) + (docWidth / 2));
            contentStream.setFont(regularFont, 16);
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18);
            contentStream.showText(String.format("Work Phase: %s", workPhase));
            contentStream.endText();
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 2);
            contentStream.showText(String.format("Percent Complete: %s%%", percentComplete));
            contentStream.endText();
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 3);
            contentStream.showText(String.format("Funded FY: %s", fundFY));
            contentStream.endText();
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 4);
            contentStream.showText(String.format("Start Date: %s", startDate));
            contentStream.endText();
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 5);
            contentStream.showText(String.format("Estimated End Date: %s", endDate));
            contentStream.endText();

            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 6);
            contentStream.showText(String.format("Architect: %s", architect));
            contentStream.endText();

            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 7);
            contentStream.showText(String.format("Engineer: %s", engineer));
            contentStream.endText();

            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart, infoBoxYStart - 18 * 8);
            contentStream.showText(String.format("Contractor: %s", contractor));
            contentStream.endText();

            // costCenter Start Box
            contentStream.setFont(boldFont, 20);
            contentStream.beginText();
            contentStream.newLineAtOffset(infoBoxXStart - 90, 200);
            contentStream.showText(String.format("Cost Center: %s", costCenter));
            contentStream.endText();

            // IO & Budget Table
            float tableYStart = 180; // Start below Comments section
            float tableXStart = margin + (docWidth / 2) - 200; // Right half of the page
            float rowHeight = 20;
            float[] columnWidths = {115, 135, 135, 135}; // Widths for I/O, Total Budget, Budget Used, Budget Remaining
            float tableWidth = columnWidths[0] + columnWidths[1] + columnWidths[2] + columnWidths[3];

            // Draw table header
            String[] headers = {"I/O", "Total Budget", "Budget Used", "Budget Remaining"};
            contentStream.setNonStrokingColor(new Color(200, 200, 200)); // Light gray for header background
            contentStream.addRect(tableXStart, tableYStart - rowHeight, tableWidth, rowHeight);
            contentStream.fill();
            contentStream.setNonStrokingColor(Color.BLACK);
            contentStream.setLineWidth(1f);
            drawTableGrid(contentStream, tableXStart, tableYStart, rowHeight, columnWidths, 2); // 2 rows (header + data)

            // Draw header text
            contentStream.setFont(boldFont, 14);
            float textX = tableXStart + 5; // Small padding inside cell
            for (int i = 0; i < headers.length; i++) {
                contentStream.beginText();
                contentStream.newLineAtOffset(textX, tableYStart - rowHeight + 4); // Center vertically
                contentStream.showText(headers[i]);
                contentStream.endText();
                textX += columnWidths[i];
            }

            // Draw data row
            //TODO: handle multiple IOs.
            String[] data = {
                    String.valueOf(internalOrder),
                    String.format(formatBudget(budgetTotal)),
                    String.format(formatBudget(budgetUsed)),
                    String.format(formatBudget(budgetRemaining))
            };

            contentStream.setFont(regularFont, 12);
            textX = tableXStart + 5;
            for (int i = 0; i < data.length; i++) {
                contentStream.beginText();
                contentStream.newLineAtOffset(textX, tableYStart - 2 * rowHeight + 4);
                contentStream.showText(data[i]);
                contentStream.endText();
                textX += columnWidths[i];
            }

            //Footer (Page Number)
            contentStream.setFont(regularFont, 10);
            contentStream.beginText();
            contentStream.newLineAtOffset(750, 20);
            contentStream.showText("Page " + document.getNumberOfPages());
            contentStream.endText();

        } catch (IOException e) {
            statusLabel.setText("Error: Failed to process file.");
            logger.error("Error processing file", e);
        }
    }

    private String formatBudget(float budgetItem) {
        String item = String.format("%.2f", budgetItem);
        StringBuilder formattedItem = new StringBuilder();

        for (int i = item.length() - 1; i > -1; i--) {
            if (i == 0) {
                formattedItem.insert(0, "$" + item.charAt(i));
            }
            else if (i == item.length() - 6 || i == item.length() - 9 || i == item.length() - 12) {
                formattedItem.insert(0, "," + item.charAt(i));
            }
            else {
                formattedItem.insert(0, item.charAt(i));
            }
        }

        return formattedItem.toString();
    }

    private void wrapText(PDPageContentStream contentStream, PDType1Font regularFont, float X, float Y, float font, int docWidth, String paragraph) throws IOException {
        float currentX = X; // Starting X position
        float currentY = Y; // Starting Y position
        float lineHeight = 16;
        contentStream.setFont(regularFont, 14);
        contentStream.beginText();
        contentStream.newLineAtOffset(currentX, currentY);
        String[] words = paragraph.split(" ");
        for (String word : words) {
            float wordWidth = font * regularFont.getStringWidth(word + " ") / 1000; // Width in points
            if (currentX + wordWidth > (docWidth/2) + 30) { // Check if word exceeds page width
                contentStream.endText();
                currentY -= lineHeight; // Move to next line
                contentStream.beginText();
                contentStream.newLineAtOffset(X, currentY);
                currentX = X;
            }
            contentStream.showText(word + " ");
            currentX += wordWidth;
        }
        contentStream.endText();
    }

    private void drawTableGrid(PDPageContentStream contentStream, float xStart, float yStart, float rowHeight, float[] columnWidths, int rowCount) throws IOException {
        contentStream.setStrokingColor(Color.BLACK);
        contentStream.setLineWidth(1f);

        // Draw horizontal lines
        for (int i = 0; i <= rowCount; i++) {
            contentStream.moveTo(xStart, yStart - i * rowHeight);
            contentStream.lineTo(xStart + sum(columnWidths), yStart - i * rowHeight);
            contentStream.stroke();
        }

        // Draw vertical lines
        float x = xStart;
        for (int i = 0; i <= columnWidths.length; i++) {
            contentStream.moveTo(x, yStart);
            contentStream.lineTo(x, yStart - rowCount * rowHeight);
            contentStream.stroke();
            if (i < columnWidths.length) {
                x += columnWidths[i];
            }
        }
    }

    private float sum(float[] array) {
        float total = 0;
        for (float v : array) {
            total += v;
        }
        return total;
    }
    private void outputFile(PDDocument document) throws IOException {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF Files", "*.pdf"));
        File outputFile = fileChooser.showSaveDialog(null);
        if (outputFile != null) {
            document.save(outputFile);
            statusLabel.setText("PDF created: " + outputFile.getName());
            logger.info("PDF created with {} pages: {}", document.getNumberOfPages(), outputFile.getName());
        } else {
            statusLabel.setText("PDF creation canceled!");
        }
    }
}