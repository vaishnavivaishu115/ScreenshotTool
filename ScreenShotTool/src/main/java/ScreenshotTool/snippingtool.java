package sniptool;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.geometry.Pos;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;


import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.prefs.Preferences;

import javax.imageio.ImageIO;

public class snippingtool extends Application {
    private XWPFDocument document;
    private AtomicInteger screenshotCounter = new AtomicInteger(0);
    private Preferences prefs;
    private TextField fileNameField;
    private TextField filePathField;
    private TextField testDescriptionField;
    private Button btnSetDefaultPath;
    private String defaultPath = ""; // Declare defaultPath with an empty string

    public snippingtool() {
        prefs = Preferences.userNodeForPackage(snippingtool.class);
        // Retrieve the default path from the preferences or use the initial default
        String initialDefaultPath = "C:\\Users\\Default\\Screenshots";
        defaultPath = prefs.get("defaultPath", initialDefaultPath);
    }

    @Override
    public void start(Stage primaryStage) {
        fileNameField = new TextField();
        fileNameField.setPromptText("File Name");
        filePathField = new TextField();
        filePathField.setPromptText("File Path");
        testDescriptionField = new TextField();
        testDescriptionField.setPromptText("Test Description");

        btnSetDefaultPath = new Button("Set as Default Path");
        btnSetDefaultPath.setOnAction(event -> {
            defaultPath = filePathField.getText();
            prefs.put("defaultPath", defaultPath);
        });

        Button btnTakeScreenshot = new Button("Take Screenshot");
        btnTakeScreenshot.setOnAction(event -> {
            // Minimize the window
            primaryStage.setIconified(true);

            // Use a separate thread to avoid UI freeze
            new Thread(() -> {
                try {
                    // Wait a short while for the window to minimize
                    Thread.sleep(500);

                    // Perform the screenshot capture on the AWT thread
                    BufferedImage screenShot = captureScreen();

                    // Add the screenshot to the Word document on the JavaFX thread
                    Platform.runLater(() -> {
                        if (document == null) {
                            prepareNewDocument();
                        }
                        addScreenshotToDocument(testDescriptionField.getText(), screenShot);
                        testDescriptionField.clear();
                        // Restore the window after the screenshot is taken
                        primaryStage.setIconified(false);
                    });
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }).start();
        });

        Button btnNewFile = new Button("New File");
        btnNewFile.setOnAction(event -> {
            prepareNewDocument();
            clearTextFields();
        });

        Button btnSaveFile = new Button("Save The File");
        btnSaveFile.setOnAction(event -> {
            String fullPath = filePathField.getText().isEmpty() ? defaultPath : filePathField.getText();
            fullPath += "\\" + fileNameField.getText() + ".docx";
            saveWordDocument(fullPath);
            clearTextFields();
            prepareNewDocument();
        });

        VBox layout = new VBox(10);
        layout.getChildren().addAll(fileNameField, filePathField, testDescriptionField, btnSetDefaultPath, btnTakeScreenshot, btnNewFile, btnSaveFile);
        layout.setPadding(new Insets(15, 20, 15, 20));
        layout.setAlignment(Pos.CENTER);

        Scene scene = new Scene(layout);
        primaryStage.setTitle("JavaFX Screenshot Tool");
        primaryStage.setScene(scene);
        primaryStage.show();
    }
    private BufferedImage captureScreen() {
        try {
            Robot robot = new Robot();
            Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
            return robot.createScreenCapture(screenRect);
        } catch (AWTException e) {
            e.printStackTrace();
            return null;
        }
    }
    private void addScreenshotToDocument(String description, BufferedImage screenShot) {
        // Assuming a page width of 6.5 inches (default Word document size) and 1-inch margins on each side
        double documentWidthInInches = 6.5;
        int maxWidthInEMU = (int) (documentWidthInInches * Units.EMU_PER_INCH);

        double imageWidthInPixels = screenShot.getWidth();
        double imageHeightInPixels = screenShot.getHeight();

        // Calculate the scale to fit the image within the page width while maintaining the aspect ratio
        double scale = maxWidthInEMU / (imageWidthInPixels * Units.EMU_PER_PIXEL);

        // Check if the image needs to be scaled down
        if (scale < 1.0) {
            imageWidthInPixels *= scale;
            imageHeightInPixels *= scale;
        }

        // Convert the scaled dimensions to EMUs
        int widthInEMU = (int) Math.round(imageWidthInPixels * Units.EMU_PER_PIXEL);
        int heightInEMU = (int) Math.round(imageHeightInPixels * Units.EMU_PER_PIXEL);

        XWPFParagraph para = document.createParagraph();
        XWPFRun runImage = para.createRun();
        runImage.setText(description);

        // Add the image to the document
        try {
            ByteArrayInputStream baos = new ByteArrayInputStream(imageToByteArray(screenShot));
            // Use the correct picture type constant from the XWPFDocument class
            int pictureType = XWPFDocument.PICTURE_TYPE_PNG;
            runImage.addPicture(baos, pictureType, "screenshot.png", widthInEMU, heightInEMU);
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }
    }




    // ... [the rest of the existing methods remain unchanged] ...
    private byte[] imageToByteArray(BufferedImage image) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    private void saveWordDocument(String filePath) {
        try (FileOutputStream out = new FileOutputStream(filePath)) {
            document.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void prepareNewDocument() {
        document = new XWPFDocument();
        screenshotCounter.set(0);
    }

    private void clearTextFields() {
        fileNameField.clear();
        filePathField.clear();
        testDescriptionField.clear();
    }

    public static void main(String[] args) {
        launch(args);
    }
}