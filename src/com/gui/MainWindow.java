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
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.text.NumberFormat;
import java.util.Vector;

//TODO: Dafür sorgen, dass bei Popups das X nicht gleichbedeutend mit OK ist. Vor allem bei Skripten!!!
public class MainWindow {
    private JButton btn_Script;
    private JComboBox cmb_Query;
    private JPanel panel_Main;
    private JTextField txtFieldStatus;
    private JTable table_Response;
    private JScrollPane pane_Table;
    private JButton btn_Import;

    public Connection connection;

    public MainWindow() {
        String str_DBPath = "jdbc:firebirdsql://localhost:3050/C:/Users/Public/Documents/MEBEDO/PROTOKOLLmanager8/DB/BackupEdit/Datenbank.FDB";
        String str_ScriptBasePath = "C:/Users/Public/Documents/Protokollmanger Advanced/";
        String str_ImportPath = "J:/Dokumentation/Betriebsmittelprüfung/Barcodes";

        //Set Style
        this.setStyle();

        //Connecte zur Datenbank
        String str_connect = connectDatabase(str_DBPath);
        if (str_connect.equals("verbunden"))
            txtFieldStatus.setText(" Erfolgreich verbunden.");
        else
            txtFieldStatus.setText(str_connect);

        this.addCheckboxItems(str_ScriptBasePath);

        btn_Script.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String str_result = executeScript(str_ScriptBasePath);
                txtFieldStatus.setText(" " + str_result);
            }
        });
        btn_Import.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                importExcel(str_ImportPath);
            }
        });
    }

    public void addCheckboxItems(String str_ScriptBasePath)
    {
        File folder = new File(str_ScriptBasePath);
        File[] listOfFiles = folder.listFiles();

        for (File file : listOfFiles) {
            if (file.isFile()) {
                cmb_Query.addItem(file.getName());
            }
        }
    }

    public void importExcel(String str_ImportPath)
    {
        txtFieldStatus.setText(" Importiere Excel-Datei");

        JFileChooser fc = new JFileChooser(str_ImportPath);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel-Datei","xlsx");
        fc.setFileFilter(filter);
        fc.setDialogTitle("Wähle die zu importierende Excel-Datei");

        int returnVal = fc.showDialog(null, "Importieren");
        File selectedFile = null;
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            selectedFile = fc.getSelectedFile();
        } else {
            txtFieldStatus.setText(" Import abgebrochen");
            return;
        }
        FileInputStream inputStream = null;
        try {
            //Open First Sheet of Excel File
            inputStream = new FileInputStream(selectedFile);
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
            DefaultTableModel returnDTM = new DefaultTableModel();
            Boolean bool_continue = true;
            while (bool_continue) {
                returnDTM = this.importDeviceFromFile(sheet.getRow(lastRow), selectedFile.getName(), dtm);
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
                    /** Erstelle PopUpBox **/
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
                    /** Ende Erstelle PopUpBox **/

                    JOptionPane.showMessageDialog(null, panel, txtFieldStatus.getText(), JOptionPane.WARNING_MESSAGE);

                    //Validiere Auswahl //Um beim nächsten weiterzumachen muss nichts verändert werden
                    if (radio2.isSelected())
                        lastRow = lastRow - Integer.parseInt(textFieldNumber.getValue().toString()) + 1; //Überspringe X (+1 weil vorher 1 schon abgezogen wurde)
                    if (radio3.isSelected())
                        bool_continue = false;
                }
            }
            workbook.close();

        } catch (FileNotFoundException ex) {
            txtFieldStatus.setText(" FileNotFoundException: " + ex.getMessage());
        } catch (IOException ex) {
            txtFieldStatus.setText(" IOException: " + ex.getMessage());
        }
    }

    public DefaultTableModel importDeviceFromFile(XSSFRow currentRow, String filename, DefaultTableModel dtm)
    {
        try {
            Vector<Object> data = new Vector<Object>();

            // Get Device and Serial Number
            String serial = "";
            String devicenumber = "";
            if (currentRow.getCell(4) != null
                    && currentRow.getCell(4).getRichStringCellValue().length() == 11) {
                devicenumber = currentRow.getCell(4).getRichStringCellValue().toString();
                serial = StringUtils.right(devicenumber, 5);
            } else {
                txtFieldStatus.setText("Die Seriennummer vom Gerät in Reihe " + (currentRow.getRowNum()+1) + " ist leer oder hat nicht die Länge 11");
                data = new Vector<Object>();
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
            String cust_id = "";
            if (resultSet.next())
                cust_id = resultSet.getString(1);
            else {
                txtFieldStatus.setText(" Kunde " + StringUtils.remove(filename, ".xlsx")
                        + " konnte nicht gefunden werden");
                data = new Vector<Object>();
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
                data = new Vector<Object>();
                data.add(serial);
                data.add(devicenumber);
                data.add("Fehler");
                data.add("SQLException beim Abfragen ob die Gerätenummer " + devicenumber + " bereits existiert");
                dtm.addRow(data);
                return dtm;
            }
            if (exists) {
                txtFieldStatus.setText(" Gerät mit der Gerätenummer " + devicenumber + " existiert bereit");
                data = new Vector<Object>();
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
            String location_id = "";
            if (resultSet.next())
                location_id = resultSet.getString(1);
            else {
                txtFieldStatus.setText(" Standort " + currentRow.getCell(5).toString()
                        + " konnte nicht gefunden werden");
                data = new Vector<Object>();
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
            String type_id = "";
            if (resultSet.next())
                type_id = resultSet.getString(1);
            else {
                data = new Vector<Object>();
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
            String dev_id = "";
            if (resultSet.next()) {
                dev_id = resultSet.getString(1);
                int conv = Integer.parseInt(dev_id);
                dev_id = Integer.toString(++conv);
            } else {
                txtFieldStatus.setText(" Fehler beim abfragen der nächsten freien DEV_ID");
                data = new Vector<Object>();
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

            data = new Vector<Object>();
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

    public Boolean checkIfExists(String serialnumber, String cust_id)
    {
        try {
            String query = "SELECT dev_id FROM device WHERE serial_no = " + serialnumber + " AND cust_id = " + cust_id + ";";
            PreparedStatement preparedStatement = connection.prepareStatement(query);
            ResultSet resultSet = preparedStatement.executeQuery();
            if (resultSet.next())
                return true;
            else
                return false;
        } catch (SQLException e) {
            txtFieldStatus.setText(" SQLException: " + e.getMessage());
            return null;
        }
    }

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
    }

    public String executeScript(String str_ScriptBasePath)
    {
        String str_ersetzen = "hierersetzen";
        try {
            String filename = cmb_Query.getSelectedItem().toString();
            Path path = Paths.get(str_ScriptBasePath + filename);

            String content = Files.readString(path, StandardCharsets.ISO_8859_1);
            if (content.contains(str_ersetzen))
            {
                /** Erstelle PopUpBox **/
                final JPanel panel = new JPanel();

                JTextPane textPane = new JTextPane();
                panel.add(textPane);

                StyledDocument doc = textPane.getStyledDocument();
                Style styleRed = textPane.addStyle("style", null);
                StyleConstants.setForeground(styleRed, Color.red);

                String[] contentSeperate = content.split(str_ersetzen);
                doc.insertString(doc.getLength(), contentSeperate[0], null); //Erste ohne Style

                for (int i = 1; i < contentSeperate.length; i++)
                {
                    doc.insertString(doc.getLength(), str_ersetzen, styleRed); //ersetzenString in Rot
                    doc.insertString(doc.getLength(), contentSeperate[i], null); //Rest ohne Farbe
                }
                /** Ende Erstelle PopUpBox **/

                JOptionPane.showMessageDialog(null, panel, "Das Skript enthält hierersetzen, was ersetzt werden muss", JOptionPane.INFORMATION_MESSAGE);

                content = textPane.getText();
                System.out.println(content);
            }

            PreparedStatement preparedStatement = null;

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
                        Vector<Object> data = new Vector<Object>();
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
                return "Es wurden " + updateCount + " Datensätze geupdatet";
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

    public static void main(String[] args) {
        //Konfiguriere und öffne MainWindow
        JFrame frame = new JFrame("MainWindow");
        frame.setContentPane(new MainWindow().panel_Main);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1000, 800);
        frame.setVisible(true);
        frame.setTitle("Protokollmanager Advanced 1.00");
    }

    public String connectDatabase(String str_DBPath)
    {
        try {
            connection = DriverManager.getConnection(
                    str_DBPath,
                    "SYSDBA", "masterkey");
            return " Verbindung erfolgreich hergestellt.";
        } catch (SQLException ex) {
            txtFieldStatus.setText("SQLException: " + ex.getMessage());
            return " SQLException: " + ex.getMessage();
        }
    }
}
