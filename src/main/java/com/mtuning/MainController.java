package com.mtuning;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Objects;
import java.util.ResourceBundle;

public class MainController implements Initializable {

    @FXML
    private Label selectedFileLabel;

    @FXML private Button selectFileButton;
    @FXML private ComboBox<String> symbolComboBox;
    @FXML private ComboBox<String> nameComboBox;
    @FXML private ComboBox<String> cdComboBox;
    @FXML private ComboBox<String> chComboBox;
    @FXML private ComboBox<String> barcodeComboBox;
    @FXML private ComboBox<String> symbol2ComboBox;
    @FXML private ComboBox<String> commentsComboBox;
    @FXML private CheckBox skipFirstRowCheckBox;
    @FXML private CheckBox editExistingProductsCheckBox;
    @FXML private Button addProductsButton;
    @FXML private TextArea logTextArea;

    private XSSFWorkbook workbook;
    ActiveXComponent oSubiekt;
    private int numberCells;
    private boolean isError;
    private String symbol;


    @Override
    public void initialize(URL location, ResourceBundle resources) {
//        ActiveXComponent oTowaryKolekcja = oSubiekt.invokeGetComponent("Towary");
//        ActiveXComponent towar = oTowaryKolekcja.invokeGetComponent("Wczytaj", new Variant("STEST1"));
//        System.out.println(towar.getProperty("Symbol"));
    }

    public void setSubiekt(ActiveXComponent oSubiekt) { // Setting the client-object in ClientViewController
        this.oSubiekt = oSubiekt;
    }

    @FXML
    public void selectFileAction(ActionEvent actionEvent) throws IOException {
        FileChooser fc = new FileChooser();
        File f = fc.showOpenDialog(null);

        if (f != null) {
            selectedFileLabel.setText(f.getAbsolutePath());
            symbolComboBox.setDisable(false);
            nameComboBox.setDisable(false);
            cdComboBox.setDisable(false);
            chComboBox.setDisable(false);
            barcodeComboBox.setDisable(false);
            symbol2ComboBox.setDisable(false);
            commentsComboBox.setDisable(false);
            skipFirstRowCheckBox.setDisable(false);
            editExistingProductsCheckBox.setDisable(false);
            addProductsButton.setDisable(false);


            try {
                FileInputStream file = new FileInputStream(f.getAbsolutePath());

                workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(0);

                // Odczytanie pierwszego wiersza z pliku Excela i przypisanie wartości jego komórek do wszystkich ComboBoxów
                Iterator<Row> rowIterator = sheet.iterator();
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();


                ObservableList<String> firstRowList = FXCollections.observableArrayList();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (cell.getCellTypeEnum().toString().equals("STRING")) {
                        firstRowList.add(cell.getStringCellValue());
                    } else {
                        firstRowList.add(String.valueOf(cell.getNumericCellValue()));
                    }
                }

                symbolComboBox.setItems(firstRowList);
                nameComboBox.setItems(firstRowList);
                cdComboBox.setItems(firstRowList);
                chComboBox.setItems(firstRowList);
                barcodeComboBox.setItems(firstRowList);
                symbol2ComboBox.setItems(firstRowList);
                commentsComboBox.setItems(firstRowList);

                file.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }
    }

    public void addProduct(ActionEvent actionEvent) {
        String col1Symbol = Objects.requireNonNull(symbolComboBox.getSelectionModel().getSelectedItem());
        String col2Name = Objects.requireNonNull(nameComboBox.getSelectionModel().getSelectedItem());
        String col3RetailPrice = Objects.requireNonNull(cdComboBox.getSelectionModel().getSelectedItem());
        String col4WholesalePrice = Objects.requireNonNull(chComboBox.getSelectionModel().getSelectedItem());
        String col5Barcode = Objects.requireNonNull(barcodeComboBox.getSelectionModel().getSelectedItem());
        String col6Symbol2 = Objects.requireNonNull(symbol2ComboBox.getSelectionModel().getSelectedItem());
        String col7Comments = Objects.requireNonNull(commentsComboBox.getSelectionModel().getSelectedItem());

        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        HashMap<Integer, Integer> columns = new HashMap<>();
        while (rowIterator.hasNext()) {
            ActiveXComponent oTowaryKolekcja = oSubiekt.invokeGetComponent("Towary");

            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                numberCells = row.getPhysicalNumberOfCells();

                // Mapowanie kolumn
                for (int i = 0; i < numberCells; i++) {
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        if (cell.toString().equals(col1Symbol)) {
                            columns.put(0, cell.getColumnIndex());
                        } else if (cell.toString().equals(col2Name)) {
                            columns.put(1, cell.getColumnIndex());
                        } else if (cell.toString().equals(col3RetailPrice)) {
                            columns.put(2, cell.getColumnIndex());
                        } else if (cell.toString().equals(col4WholesalePrice)) {
                            columns.put(3, cell.getColumnIndex());
                        } else if (cell.toString().equals(col5Barcode)) {
                            columns.put(4, cell.getColumnIndex());
                        } else if (cell.toString().equals(col6Symbol2)) {
                            columns.put(5, cell.getColumnIndex());
                        } else if (cell.toString().equals(col7Comments)) {
                            columns.put(6, cell.getColumnIndex());
                        }
                    }
                }
            }

            if (skipFirstRowCheckBox.isSelected() && row.getRowNum() == 0) {
                continue;
            } else {
                isError = false;
                numberCells = row.getPhysicalNumberOfCells();
                ActiveXComponent towar;

                Variant czyIstnieje = oTowaryKolekcja.invoke("Istnieje", row.getCell(columns.get(0)).toString());
                if (czyIstnieje.getBoolean()) {
                    towar = oTowaryKolekcja.invokeGetComponent("Wczytaj", new Variant(row.getCell(columns.get(0)).toString()));
                } else {
                    towar = oTowaryKolekcja.invokeGetComponent("Dodaj", new Variant("1"));
                    towar.setProperty("Symbol", row.getCell(columns.get(0)).toString());
                }

                symbol = row.getCell(columns.get(0)).toString();

                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    if (!editExistingProductsCheckBox.isSelected() && cell.getColumnIndex() == columns.get(0)) {
                        if (czyIstnieje.getBoolean()) {
                            isError = true;
                            System.out.println("Towar o symbolu " + cell.toString() + " już istnieje.");
                            logTextArea.setText(logTextArea.getText() + "Błąd! Towar o symbolu " + cell.toString() + " już istnieje.\n");
                            break;
                        }
                    } else if (cell.getColumnIndex() == columns.get(1)) {
                        int stringSize = cell.toString().length();

                        if (stringSize <= 50) {
                            towar.setProperty("Nazwa", cell.toString());
                        } else {
                            String correctedName = cell.toString().substring(0, 50);
                            towar.setProperty("Nazwa", correctedName);
                        }

                        towar.setProperty("Opis", cell.toString());
                    } else if (cell.getColumnIndex() == columns.get(2)) {
                        if (cell.getCellTypeEnum().toString().equals("NUMERIC")) {
                            ActiveXComponent cenyTowaru = towar.invokeGetComponent("Ceny");
                            ActiveXComponent oCenaDetaliczna = cenyTowaru.invokeGetComponent("Element", new Variant(1));
                            String cenaDetaliczna = cell.toString();
                            cenaDetaliczna = cenaDetaliczna.replace('.',',');
                            oCenaDetaliczna.setProperty("Brutto", cenaDetaliczna);
                        } else {
                            isError = true;
                            System.out.println("Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie ceny detalicznej.");
                            logTextArea.setText(logTextArea.getText() + "Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie ceny detalicznej.\n");
                            break;
                        }
                    } else if (cell.getColumnIndex() == columns.get(3)) {
                        if (cell.getCellTypeEnum().toString().equals("NUMERIC")) {
                            ActiveXComponent cenyTowaru = towar.invokeGetComponent("Ceny");
                            ActiveXComponent oCenaHurtowa = cenyTowaru.invokeGetComponent("Element", new Variant(2));
                            String cenaHurtowa = cell.toString();
                            cenaHurtowa = cenaHurtowa.replace('.',',');
                            oCenaHurtowa.setProperty("Brutto", cenaHurtowa);
                        } else {
                            isError = true;
                            System.out.println("Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie ceny hurtowej.");
                            logTextArea.setText(logTextArea.getText() + "Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie ceny hurtowej.\n");
                            break;
                        }
                    } else if (cell.getColumnIndex() == columns.get(4)) {
                        if (cell.getCellTypeEnum().toString().equals("NUMERIC")) {
                            ActiveXComponent kodyKreskowe = towar.invokeGetComponent("KodyKreskowe");
                            kodyKreskowe.setProperty("Podstawowy", new Variant(cell.getNumericCellValue()));
                        } else {
                            isError = true;
                            System.out.println("Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie kodu kreskowego.");
                            logTextArea.setText(logTextArea.getText() + "Nie można dodać towaru z wiersza " + (row.getRowNum() + 1) + " o symbolu " + symbol + ". Nieprawidłowa wartość - " + cell.toString() + " w kolumnie kodu kreskowego.\n");
                            break;
                        }
                    } else if (cell.getColumnIndex() == columns.get(5)) {
                        towar.setProperty("SymbolUDostawcy", cell.toString());
                    } else if (cell.getColumnIndex() == columns.get(6)) {
                        towar.setProperty("Uwagi", cell.toString());
                    }

                    if ((towar.getProperty("Symbol") + " " + towar.getProperty("Nazwa")).length() > 40) {
                        towar.setProperty("NazwaDlaUF", ((towar.getProperty("Symbol") + " " + towar.getProperty("Nazwa"))).substring(0, 40));
                    } else {
                        towar.setProperty("NazwaDlaUF", ((towar.getProperty("Symbol") + " " + towar.getProperty("Nazwa"))));
                    }
                }

                if (!isError) {
                    towar.invoke("Zapisz");
                    towar.invoke("Zamknij");
                }
            }
        }
    }
}
