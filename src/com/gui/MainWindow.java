package com.gui;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.text.BadLocationException;
import javax.swing.text.Style;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.text.NumberFormat;
import java.util.Vector;

public class MainWindow {
    private JButton btn_Script;
    private JComboBox<String> cmb_Query;
    private JPanel panel_Main;
    private JTextField txtFieldStatus;
    private JTable table_Response;
    private JScrollPane pane_Table;
    private JButton btn_Import;
    private static final String VERSION = "1.04";
    private static final String ERSETZEN = "hierersetzen";

    public Connection connection;

    /** Programm Start **/
    public static void main(String[] args) {
        //Konfiguriere und öffne MainWindow
        JFrame frame = new JFrame("MainWindow");
        frame.setContentPane(new MainWindow().panel_Main);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1000, 800);
        frame.setVisible(true);
        frame.setTitle("Protokollmanager Advanced " + VERSION);
    }

    /** Programm Start **/
    public MainWindow() {
        //Set Style
        this.setStyle();

        //Read Paths
        String[] paths = new String[3];
        paths = this.readPaths(paths);
        if (paths == null)
            return;
        String str_DBPath = paths[0];
        String str_ScriptBasePath = paths[1] + "/";
        String str_ImportPath = paths[2] + "/";

        //Connecte zur Datenbank
        connectDatabase(str_DBPath);

        this.addCheckboxItems(str_ScriptBasePath);

        btn_Script.addActionListener(e -> {
            String str_result = executeScript(str_ScriptBasePath);
            txtFieldStatus.setText(" " + str_result);
        });
        btn_Import.addActionListener(e -> importExcel(str_ImportPath));
    }

    /** Programm Start Hilfsmethode**/
    public void setStyle() {
        Color color_Background = new Color(32, 136, 203);
        table_Response.setBackground(color_Background);
        panel_Main.setBackground(color_Background);
        txtFieldStatus.setBackground(color_Background);
        txtFieldStatus.setForeground(Color.white);
        txtFieldStatus.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        table_Response.setForeground(Color.white);
        table_Response.setSelectionBackground(Color.red);
        table_Response.setSelectionForeground(Color.white);
        pane_Table.getViewport().setBackground(Color.white);
        pane_Table.setBorder(javax.swing.BorderFactory.createLineBorder(Color.white));
        table_Response.setGridColor(Color.white);
        table_Response.setAutoCreateRowSorter(true);

        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } //Setze Windows Look!
        catch (UnsupportedLookAndFeelException e) {
            JOptionPane.showMessageDialog(null,
                    "Fehler beim Setzen des Windows Designs: " + e.getMessage(),
                    "UnsupportedLookAndFeelException",
                    JOptionPane.ERROR_MESSAGE);
        }
        catch (ClassNotFoundException e) {
            JOptionPane.showMessageDialog(null,
                    "Fehler beim Setzen des Windows Designs: " + e.getMessage(),
                    "ClassNotFoundException",
                    JOptionPane.ERROR_MESSAGE);
        }
        catch (InstantiationException e) {
            JOptionPane.showMessageDialog(null,
                    "Fehler beim Setzen des Windows Designs: " + e.getMessage(),
                    "InstantiationException",
                    JOptionPane.ERROR_MESSAGE);
        }
        catch (IllegalAccessException e) {
            JOptionPane.showMessageDialog(null,
                    "Fehler beim Setzen des Windows Designs: " + e.getMessage(),
                    "IllegalAccessException",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Pfade einlesen falls paths.cfg vorhanden **/
    public String[] readPaths(String[] paths)
    {
        File config = new File(System.getenv("PUBLIC") + "/Documents/Protokollmanger Advanced/paths.cfg");
        String error = "";
        if (config.exists()) {
            try {
                FileReader fr = new FileReader(config, StandardCharsets.UTF_8);
                BufferedReader br = new BufferedReader(fr);
                String line;
                int i = 0;
                while ((line = br.readLine()) != null) {
                    paths[i] = line;
                    i++;
                }
                fr.close();
                if (paths[0] == null) {
                    error += "Datenbankpfad ist null! ";
                } else if (checkIfFileOrDirectoryDoesntExist(paths[0])) {
                    error += "Datenbankpfad " + paths[0] + " ist ungültig! ";
                    paths[0] = null;
                }
                if (paths[1] == null) {
                    error += "Skriptpfad ist null! ";
                } else if (checkIfFileOrDirectoryDoesntExist(paths[1])) {
                    error += "Skriptpfad " + paths[1] + " ist ungültig! ";
                    paths[1] = null;
                }
                if (paths[2] == null) {
                    error += "Excelpfad ist null! ";
                } else if (checkIfFileOrDirectoryDoesntExist(paths[2])) {
                    error += "Excelpfad " + paths[2] + " ist ungültig! ";
                    paths[2] = null;
                }
            } catch (FileNotFoundException e) {
                error = "FileNotFoundException beim öffnen von Öffentliche Dokumente/Protokollmanager Advanced/paths.cfg!";
            } catch (IOException e) {
                error = "IOException beim öffnen von Öffentliche Dokumente/Protokollmanager Advanced/paths.cfg!";
            }
            if (paths[0] == null || paths[1] == null || paths[2] == null){
                txtFieldStatus.setText(" "+ error);
                this.disableAll(true);

                JOptionPane.showMessageDialog(null,
                        error + "Bitte Programm neustarten und gegebenenfalls paths.cfg korrigieren oder löschen",
                        "Fehler beim Import von paths.cfg",
                        JOptionPane.ERROR_MESSAGE);

                return null;
            }
        }
        else {
            paths = this.setPathsOnStart();
            if (paths[0] == null) {
                error += " Datenbankpfad ist null!";
            } else {
                paths[0] = "jdbc:firebirdsql://localhost:3050/" + paths[0]; //local firebirdsql server
            }
            if (paths[1] == null) {
                error += " Skriptpfad ist null!";
            }
            if (paths[2] == null) {
                error += " Excelpfad ist null!";
            }
            if (paths[0] == null || paths[1] == null || paths[2] == null){
                txtFieldStatus.setText(" "+ error);
                this.disableAll(true);

                JOptionPane.showMessageDialog(null,
                        error + " Bitte Programm neustarten und gültige Pfade angeben",
                        "Fehler beim Setzen der Pfade",
                        JOptionPane.ERROR_MESSAGE);

                return null;
            }
            if (!this.savePaths(paths))
                JOptionPane.showMessageDialog(null,
                        "Pfade konnten nicht in Öffentliche Dokumente/Protokollmanager Advanced/paths.cfg gespeichert werden" +
                                " und müssen beim nächsten Start erneut eingegeben werden.",
                        "IOException",
                        JOptionPane.WARNING_MESSAGE);
        }

        return paths;
    }

    /** Testen ob Pfade korrekt sind **/
    public boolean checkIfFileOrDirectoryDoesntExist(String path)
    {
        if (path.startsWith("jdbc:firebirdsql://localhost:3050/"))
            path = StringUtils.remove(path, "jdbc:firebirdsql://localhost:3050/");
        File tmpDir = new File(path);
        return !tmpDir.exists();
    }

    /** Alle Button disablen, falls ein Pfad null oder ungültig ist **/
    public void disableAll(boolean disable)
    {
        btn_Import.setEnabled(disable);
        btn_Script.setEnabled(disable);
        cmb_Query.setEnabled(disable);
    }

    /** Ansonsten Pfade setzen lassen **/
    public String[] setPathsOnStart()
    {
        JOptionPane.showMessageDialog(null,
                "Die Datei paths.cfg, in welcher Pfade gespeichert werden, ist noch nicht vorhanden. Deswegen müssen jetzt 3 Pfade gesetzt werden.",
                "paths.cfg ist noch nicht vorhanden",
                JOptionPane.INFORMATION_MESSAGE);

        String[] paths = new String[3];

        paths[0] = this.fileDialogOnStart("Setze den Pfad für Datenbank.FDB","Datenbank.FDB","fdb", false);
        paths[1] = this.fileDialogOnStart("Setze den Pfad für den Skriptordner","Skriptordner","sql",true);
        paths[2] = this.fileDialogOnStart("Setze den Pfad für die Exceldateien","Exceldateien","xlsx", true);

        return paths;
    }

    /** Ansonsten Pfade setzen lassen Hilfsmethode**/
    public String fileDialogOnStart(String status, String filtername, String filtertype, boolean directory)
    {
        txtFieldStatus.setText(" " + status);

        JFileChooser fc;
        if (directory) {
            fc = new JFileChooser(System.getenv("PUBLIC")+"/Documents") {
                public void approveSelection() {
                    if (!getSelectedFile().isFile())
                        super.approveSelection();
                }
            };
            fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        } else {
            fc = new JFileChooser(System.getenv("PUBLIC")+"/Documents");
            fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
        }

        FileNameExtensionFilter filter = new FileNameExtensionFilter(filtername,filtertype);
        fc.setFileFilter(filter);
        fc.setDialogTitle(status);

        int returnVal = fc.showDialog(null, "Auswählen");
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File selection = fc.getSelectedFile();
            return selection.getAbsolutePath();
        } else {
            return null;
        }
    }

    /** Gesetzte Pfade in paths.cfg speichern **/
    public boolean savePaths(String[] paths)
    {
        try {
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(
                    System.getenv("PUBLIC")+"/Documents/Protokollmanger Advanced/paths.cfg"), StandardCharsets.UTF_8));
            writer.write(paths[0]);
            writer.newLine();
            writer.write(paths[1]);
            writer.newLine();
            writer.write(paths[2]);
            writer.close();

            return true;
        } catch (IOException e) {
            txtFieldStatus.setText(" IOException beim Speichern der Pfade: " + e.getMessage());
            return false;
        }
    }

    /** Items zur Checkbox hinzufügen **/
    public void addCheckboxItems(String str_ScriptBasePath)
    {
        File folder = new File(str_ScriptBasePath);

        File[] listOfFiles = folder.listFiles();

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile() && file.getName().toLowerCase().endsWith(".sql")) {
                    cmb_Query.addItem(file.getName());
                }
            }
        }

        if (cmb_Query.getItemCount() == 0)
        {
            cmb_Query.addItem("Keine .sql Dateien vorhanden! " + str_ScriptBasePath);
            cmb_Query.setEnabled(false);
            btn_Script.setEnabled(false);
        }
    }

    /** Zur Datenbank connecten **/
    public void connectDatabase(String str_DBPath)
    {
        if (checkIfFileOrDirectoryDoesntExist(str_DBPath))
        {
            txtFieldStatus.setText(" Die Datenbank befindet sich nicht mehr in: " + str_DBPath);
            btn_Script.setEnabled(false);
            btn_Import.setEnabled(false);
            return;
        }

        try {
            connection = DriverManager.getConnection(
                    str_DBPath,
                    "SYSDBA", "masterkey");
            txtFieldStatus.setText(" Verbindung erfolgreich hergestellt.");
        } catch (SQLException ex) {
            txtFieldStatus.setText(" SQLException: " + ex.getMessage());
        }
    }

    /** Importiere aus einer Excel Datei **/
    public void importExcel(String str_ImportPath)
    {
        txtFieldStatus.setText(" Importiere Excel-Datei");

        File folder = new File(str_ImportPath);
        if (!folder.exists()) {
            JOptionPane.showMessageDialog(null,
                    "Der Pfad für den Excelordner ist ungültig! Daher wird der FileDialog im Dokumente Ordner des derzeitigen Nutzers starten.",
                    "Ungültiger Pfad! " + str_ImportPath,
                    JOptionPane.WARNING_MESSAGE);
        }

        JFileChooser fc = new JFileChooser(str_ImportPath);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel-Datei","xlsx");
        fc.setFileFilter(filter);
        fc.setDialogTitle("Wähle die zu importierende Excel-Datei");

        int returnVal = fc.showDialog(null, "Importieren");
        File selectedFile;
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            selectedFile = fc.getSelectedFile();
        } else {
            txtFieldStatus.setText(" Import abgebrochen");
            return;
        }
        try {
            //Open First Sheet of Excel File
            FileInputStream inputStream = new FileInputStream(selectedFile);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            //Get Last Row
            int lastRow = sheet.getLastRowNum();
            //Go 1 Row up until there is a row with a . at the 10th cell
            while (sheet.getRow(lastRow).getCell(10) == null || // Zeile darf nicht null sein
                    !sheet.getRow(lastRow).getCell(10).toString().contains("-")) // - da das Datumsformat zu 1-Januar-2020 umformatiert wird
            {
                lastRow--;
            }

            DefaultTableModel dtm = new DefaultTableModel() {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            dtm.addColumn("Seriennummer");
            dtm.addColumn("Gerätenummer");
            dtm.addColumn("Status");
            dtm.addColumn("Protokoll");
            DefaultTableModel returnDTM;
            boolean bool_continue = true;
            int countImport = 0;
            while (bool_continue) {
                returnDTM = this.importDeviceFromFile(sheet.getRow(lastRow), selectedFile.getName(), dtm);
                if (returnDTM.getValueAt(returnDTM.getRowCount()-1,3).toString().contains("erfolgreich")) //Zähle Imports mit
                    countImport++;
                lastRow--;

                //Aktualisiere Tabel Inhalt
                dtm = returnDTM;
                table_Response.setModel(dtm);
                table_Response.getColumn(table_Response.getColumnName(0)).setMaxWidth(105);         //0 = Seriennummer
                table_Response.getColumn(table_Response.getColumnName(0)).setPreferredWidth(105);
                table_Response.getColumn(table_Response.getColumnName(1)).setMaxWidth(105);         //1 = Gerätenummer
                table_Response.getColumn(table_Response.getColumnName(1)).setPreferredWidth(105);
                table_Response.getColumn(table_Response.getColumnName(2)).setMaxWidth(105);         //2 = Status
                table_Response.getColumn(table_Response.getColumnName(2)).setPreferredWidth(105);

                if (returnDTM.getValueAt(returnDTM.getRowCount()-1,2).toString().equals("Fehler")) //Wenn der letzte Eintrag "Fehler" enthält
                {
                    //Erstelle PopUpBox
                    final JPanel panel = new JPanel();
                    final JRadioButton radio1 = new JRadioButton("Weiter");
                    final JRadioButton radio2 = new JRadioButton("X überspringen und Weiter, X =");
                    final JRadioButton radio3 = new JRadioButton("Abbrechen");

                    NumberFormat amountFormat = NumberFormat.getNumberInstance();
                    JFormattedTextField textFieldNumber = new JFormattedTextField(amountFormat);
                    textFieldNumber.setValue(0); //DefaultValue wird immer benutzt, wenn ungültige Eingabe vorliegt, deswegen auf 0
                    textFieldNumber.setColumns(2);

                    ButtonGroup G = new ButtonGroup();
                    G.add(radio1);
                    G.add(radio2);
                    G.add(radio3);
                    panel.add(radio1);
                    panel.add(radio2);
                    panel.add(textFieldNumber);
                    panel.add(radio3);
                    radio1.setSelected(true);
                    //Ende Erstelle PopUpBox

                    int n = JOptionPane.showConfirmDialog(null, panel, txtFieldStatus.getText(), JOptionPane.DEFAULT_OPTION,JOptionPane.WARNING_MESSAGE);
                    if (n != JOptionPane.OK_OPTION) //Wenn X geklickt wurde
                        bool_continue = false;
                    else {
                        //Validiere Auswahl //Um beim nächsten weiterzumachen muss nichts verändert werden
                        if (radio2.isSelected())
                            lastRow = lastRow - Integer.parseInt(textFieldNumber.getValue().toString()) + 1; //Überspringe X (+1 weil vorher 1 schon abgezogen wurde)
                        if (radio3.isSelected())
                            bool_continue = false;
                    }
                }
            }
            workbook.close();
            txtFieldStatus.setText(" Es wurden " + countImport + " Geräte erfolgreich von " + selectedFile.getName() + " importiert.");

        } catch (FileNotFoundException ex) {
            txtFieldStatus.setText(" FileNotFoundException: " + ex.getMessage());
        } catch (IOException ex) {
            txtFieldStatus.setText(" IOException: " + ex.getMessage());
        }
    }

    /** Importiere aus einer Excel Datei Hilfsmethode**/
    public DefaultTableModel importDeviceFromFile(XSSFRow currentRow, String filename, DefaultTableModel dtm)
    {
        try {
            Vector<Object> data = new Vector<>();

            // Get Device and Serial Number
            String serial;
            String devicenumber;
            if (currentRow.getCell(4) != null
                    && currentRow.getCell(4).getRichStringCellValue().length() == 11) {
                devicenumber = currentRow.getCell(4).getRichStringCellValue().toString();
                serial = StringUtils.right(devicenumber, 5);
            } else {
                txtFieldStatus.setText("Die Seriennummer vom Gerät in Reihe " + (currentRow.getRowNum()+1) + " ist leer oder hat nicht die Länge 11");
                if (currentRow.getCell(3) != null)
                    data.add(currentRow.getCell(3).toString());
                else
                    data.add("?");
                if (currentRow.getCell(4) != null)
                    data.add(currentRow.getCell(4).getRichStringCellValue().toString());
                else
                    data.add("?");
                data.add("Fehler");
                data.add("Seriennummer vom Gerät in Reihe " + (currentRow.getRowNum()+1) + " ist leer oder hat nicht die Länge 11");
                dtm.addRow(data);
                return dtm;
            }

            // Get customer ID
            String query = "SELECT cust_id FROM customer WHERE f_acronym = '"
                    + StringUtils.remove(filename, ".xlsx") + "';";
            PreparedStatement preparedStatement = connection.prepareStatement(query);
            ResultSet resultSet = preparedStatement.executeQuery();
            String cust_id;
            if (resultSet.next())
                cust_id = resultSet.getString(1);
            else {
                txtFieldStatus.setText(" Kunde " + StringUtils.remove(filename, ".xlsx")
                        + " konnte nicht gefunden werden");
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("Kunde " + StringUtils.remove(filename, ".xlsx") + " konnte nicht gefunden werden");
                dtm.addRow(data);
                return dtm;
            }

            // Check if it already exists
            Boolean exists = this.checkIfExists(serial, cust_id);
            if (exists == null){
                txtFieldStatus.setText(" SQLException beim Abfragen ob die Gerätenummer " + devicenumber + " bereits existiert");
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("SQLException beim Abfragen ob die Gerätenummer " + devicenumber + " bereits existiert");
                dtm.addRow(data);
                return dtm;
            }
            if (exists) {
                txtFieldStatus.setText(" Gerät mit der Gerätenummer " + devicenumber + " existiert bereit");
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("Gerät mit der Gerätenummer " + devicenumber + " existiert bereit");
                dtm.addRow(data);
                return dtm;
            }

            // Get Location ID
            query = "SELECT location_id FROM location WHERE location_name LIKE '" + currentRow.getCell(5).toString() + "';";
            preparedStatement = connection.prepareStatement(query);
            resultSet = preparedStatement.executeQuery();
            String location_id;
            if (resultSet.next())
                location_id = resultSet.getString(1);
            else {
                txtFieldStatus.setText(" Standort " + currentRow.getCell(5).toString()
                        + " konnte nicht gefunden werden");
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("Standort " + currentRow.getCell(5).toString()
                        + " konnte nicht gefunden werden");
                dtm.addRow(data);
                return dtm;
            }

            // Get harzard class
            String f_hazard_class = "5";
            if (location_id.equals("172") && cust_id.equals("38")) //Carat und Werkstatt?
                f_hazard_class = "4";

            // Get Type ID
            query = "SELECT type_id FROM dev_type WHERE type_name = '" + currentRow.getCell(6).toString() + "';";
            preparedStatement = connection.prepareStatement(query);
            resultSet = preparedStatement.executeQuery();
            String type_id;
            if (resultSet.next())
                type_id = resultSet.getString(1);
            else {
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Warnung");
                data.add("Gerätetyp " + currentRow.getCell(6).toString()
                        + " konnte nicht gefunden werden und wurde auf Unbekannt gesetzt");
                dtm.addRow(data);
                type_id = "-1";
            }

            //GET DEV ID
            query = "SELECT MAX(dev_id) FROM device";
            preparedStatement = connection.prepareStatement(query);
            resultSet = preparedStatement.executeQuery();
            String dev_id;
            if (resultSet.next()) {
                dev_id = resultSet.getString(1);
                int conv = Integer.parseInt(dev_id);
                dev_id = Integer.toString(++conv);
            } else {
                txtFieldStatus.setText(" Fehler beim abfragen der nächsten freien DEV_ID");
                data = new Vector<>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("Fehler beim abfragen der nächsten freien DEV_ID");
                dtm.addRow(data);
                return dtm;
            }

            query = "INSERT INTO device ("
                    + "dev_id,"
                    + "cust_id,"
                    + "type_id,"
                    + "dev_no,"
                    + "serial_no,"
                    + "location_id,"
                    + "f_hazard_class,"
                    + "status,"
                    + "report_no_gen) VALUES ("
                    + "?,?,?,?,?,?,?,?,?)";

            PreparedStatement st = connection.prepareStatement(query);
            st.setString(1, dev_id);
            st.setString(2, cust_id);
            st.setString(3, type_id);
            st.setString(4, devicenumber);
            st.setString(5, serial);
            st.setString(6, location_id);
            st.setString(7, f_hazard_class);
            st.setString(8, "3");
            st.setString(9, "0");

            if (st.executeUpdate() != 1) {
                txtFieldStatus.setText("Beim Insert-Statement von " + devicenumber + " ist wohl ein Fehler aufgetreten.");
            }

            txtFieldStatus.setText(" " + filename + " wurde erfolgreich importiert");

            data = new Vector<>();
            data.add(serial);
            data.add(devicenumber);
            data.add("Info");
            data.add("wurde erfolgreich importiert");
            dtm.addRow(data);
            return dtm;
        }
        catch (SQLException ex) {
            txtFieldStatus.setText(" SQLException: " + ex.getMessage());
            return null;
        }
    }

    /** Teste ob ein Gerät schon in der Datenbank ist **/
    public Boolean checkIfExists(String serialnumber, String cust_id)
    {
        try {
            String query = "SELECT dev_id FROM device WHERE serial_no = " + serialnumber + " AND cust_id = " + cust_id + ";";
            PreparedStatement preparedStatement = connection.prepareStatement(query);
            ResultSet resultSet = preparedStatement.executeQuery();
            return resultSet.next();
        } catch (SQLException e) {
            txtFieldStatus.setText(" SQLException: " + e.getMessage());
            return null;
        }
    }

    /** Führe ein Skript aus **/
    public String executeScript(String str_ScriptBasePath)
    {
        if (cmb_Query.getSelectedItem() != null)
            txtFieldStatus.setText(" Führe " + cmb_Query.getSelectedItem().toString() + " aus...");
        else
            return " Fehler! Das aktive Element der Combobox ist null!";

        if (checkIfFileOrDirectoryDoesntExist(str_ScriptBasePath))
        {
            return " Die Datei " + str_ScriptBasePath + " existiert nicht!";
        }

        try {
            String filename = cmb_Query.getSelectedItem().toString();
            Path path = Paths.get(str_ScriptBasePath + filename);

            String content = Files.readString(path, StandardCharsets.UTF_8);
            if (content.contains(ERSETZEN))
            {
                //Erstelle PopUpBox
                final JPanel panel = new JPanel();

                JTextPane textPane = new JTextPane();
                panel.add(textPane);

                StyledDocument doc = textPane.getStyledDocument();
                Style styleRed = textPane.addStyle("style", null);
                StyleConstants.setForeground(styleRed, Color.red);

                String[] contentSeperate = content.split(ERSETZEN);
                doc.insertString(doc.getLength(), contentSeperate[0], null); //Erste ohne Style

                for (int i = 1; i < contentSeperate.length; i++)
                {
                    doc.insertString(doc.getLength(), ERSETZEN, styleRed); //ersetzenString in Rot
                    doc.insertString(doc.getLength(), contentSeperate[i], null); //Rest ohne Farbe
                }
                //Ende Erstelle PopUpBox

                int n = JOptionPane.showConfirmDialog(null, panel, "Das Skript enthält " + ERSETZEN + ", was ersetzt werden muss", JOptionPane.DEFAULT_OPTION, JOptionPane.INFORMATION_MESSAGE);
                if (n != JOptionPane.OK_OPTION) //Wenn X geklickt wurde
                    return " Ausführung von " + cmb_Query.getSelectedItem().toString() + " wurde abgebrochen!";

                n = JOptionPane.showConfirmDialog(
                        null,
                        doc.getText(0,doc.getLength()),
                        "Soll das Skript so abgesendet werden?",
                        JOptionPane.YES_NO_OPTION);

                if (n != JOptionPane.YES_OPTION) //Wenn No oder X geklickt wurde
                    return " Ausführung von " + cmb_Query.getSelectedItem().toString() + " wurde abgebrochen!";

                content = textPane.getText();
            }

            PreparedStatement preparedStatement;

            if (filename.startsWith("Check") || filename.startsWith("Get")) {
                preparedStatement = connection.prepareStatement(content);
                try (ResultSet resultSet = preparedStatement.executeQuery()) {
                    ResultSetMetaData rsmd = resultSet.getMetaData();
                    int columnCount = rsmd.getColumnCount();
                    DefaultTableModel dtm = new DefaultTableModel() {
                        @Override
                        public boolean isCellEditable(int row, int column) {
                            return false;
                        }
                    };

                    for (int i = 1; i <= columnCount; i++) {
                        dtm.addColumn(rsmd.getColumnName(i));
                    }
                    while (resultSet.next()) {
                        Vector<Object> data = new Vector<>();
                        for (int i = 1; i <= columnCount; i++) {
                            data.add(resultSet.getString(i));

                        }
                        dtm.addRow(data);
                    }
                    table_Response.setModel(dtm);
                }
            }
            if (filename.startsWith("Set") || filename.startsWith("Update"))
            {
                table_Response.setModel(new DefaultTableModel()); // Remove Tablecontent

                //updateCount ist sowohl zum zählen der update statements als auch nachher zum wiedergeben der geupdateten zeilen
                int updateCount = StringUtils.countMatches(content.toLowerCase(), "update");
                if (updateCount > 1)
                {
                    String[] contentSeperate = content.toLowerCase().split("update");
                    updateCount = 0;
                    for (int i = 1; i < contentSeperate.length; i++)
                    {
                        preparedStatement = connection.prepareStatement("update" + contentSeperate[i]);
                        updateCount = updateCount + preparedStatement.executeUpdate();
                    }
                }
                else {
                    preparedStatement = connection.prepareStatement(content);
                    updateCount = preparedStatement.executeUpdate();
                }
                return "Es wurden " + updateCount + " Datensätze geupdatet durch " + cmb_Query.getSelectedItem().toString();
            }

            return "Skript " + cmb_Query.getSelectedItem().toString() + " wurde erfolgreich ausgeführt.";
        } catch (FileNotFoundException e) {
            return "FileNotFoundException: " + e.getMessage();
        } catch (IOException e) {
            return "IOException: " + e.getMessage();
        } catch (SQLException e) {
            return "SQLException: " + e.getMessage();
        } catch (BadLocationException e) {
            return "BadLocationException" + e.getMessage();
        }
    }
}
