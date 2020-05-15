package com.gui;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.sound.midi.SysexMessage;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.util.Vector;

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
            while (!sheet.getRow(lastRow).getCell(10).toString().contains("-")) // - da das Datumsformat zu 1-Januar-2020 umformatiert wird
            {
                lastRow--;
            }
            XSSFRow currentRow = sheet.getRow(lastRow);

            // Get Device and Serial Number
            if (currentRow.getCell(1).toString().length() < 4 || currentRow.getCell(2).toString().length() < 2
                    || currentRow.getCell(3).toString().length() < 5)
            {
                txtFieldStatus.setText(" Ungültige Seriennummer in Zeile " + lastRow + ": " + currentRow.getCell(1).toString()
                        + currentRow.getCell(2).toString() + currentRow.getCell(3).toString());
                return;
            }
            String barcode = StringUtils.left(currentRow.getCell(1).toString(),4);
            String year = StringUtils.left(currentRow.getCell(2).toString(),2);
            String serial = StringUtils.left(currentRow.getCell(3).toString(),5);
            String devicenumber = barcode + year + serial;

            // Get customer ID
            String query = "SELECT cust_id FROM customer WHERE f_acronym = '"
                    + StringUtils.remove(selectedFile.getName(),".xlsx") + "';";
            PreparedStatement preparedStatement = connection.prepareStatement(query);
            ResultSet resultSet = preparedStatement.executeQuery();
            String cust_id = "";
            if (resultSet.next())
                cust_id = resultSet.getString(1);
            else
            {
                 txtFieldStatus.setText(" Kunde " + StringUtils.remove(selectedFile.getName(),".xlsx")
                         + " konnte nicht gefunden werden");
                 return;
            }

            // Check if it already exists
            Boolean exists = this.checkIfExists(serial, cust_id);
            if (exists == null)
                return;
            if (exists)
            {
                txtFieldStatus.setText(" Gerät mit der Gerätenummer " + devicenumber + " existiert bereit");
                return;
            }

            // Get Location ID
            query = "SELECT location_id FROM location WHERE location_name LIKE '" + currentRow.getCell(5).toString() + "';";
            preparedStatement = connection.prepareStatement(query);
            resultSet = preparedStatement.executeQuery();
            String location_id = "";
            if (resultSet.next())
                location_id = resultSet.getString(1);
            else
            {
                txtFieldStatus.setText(" Standort " + currentRow.getCell(5).toString()
                        + " konnte nicht gefunden werden");
                return;
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
            else
            {
                txtFieldStatus.setText(" Gerätytyp " + currentRow.getCell(6).toString()
                        + " konnte nicht gefunden werden");
                return;
            }

            query = "INSERT INTO device ("
                    + "cust_id,"
                    + "type_id,"
                    + "dev_no,"
                    + "serial_no,"
                    + "location_id,"
                    + "f_hazard_class,"
                    + "status,"
                    + "report_no_gen) VALUES ("
                    + "?,?,?,?,?,?,?,?)";

            PreparedStatement st = connection.prepareStatement(query);
            st.setString(1, cust_id);
            st.setString(2, type_id);
            st.setString(3, devicenumber);
            st.setString(4, serial);
            st.setString(5, location_id);
            st.setString(6, f_hazard_class);
            st.setString(7, "3");
            st.setString(8, "0");

            //TODO: PreparedStatement ausführen!

            workbook.close();
            txtFieldStatus.setText(" " + selectedFile.getName() + " wurde erfolgreich importiert");
        } catch (FileNotFoundException ex) {
            txtFieldStatus.setText(" FileNotFoundException: " + ex.getMessage());
        } catch (IOException ex) {
            txtFieldStatus.setText(" IOException: " + ex.getMessage());
        } catch (SQLException ex) {
            txtFieldStatus.setText(" SQLException: " + ex.getMessage());
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
        try {
            //TODO: SetRoom.sql Parameter ersetzen, SetAllHazardClass.sql wirft bei 2. UPDATE Fehler aus

            String filename = cmb_Query.getSelectedItem().toString();
            Path path = Paths.get(str_ScriptBasePath + filename);
            String content = Files.readString(path, StandardCharsets.ISO_8859_1);
            System.out.println(content);
            PreparedStatement preparedStatement = connection.prepareStatement(content);

            if (filename.startsWith("Check") || filename.startsWith("Get")) {
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
                int updateCount = preparedStatement.executeUpdate();
                return "Es wurden " + updateCount + " Datensätze geupdatet";
            }

            return "Skript " + cmb_Query.getSelectedItem().toString() + " wurde erfolgreich ausgeführt.";
        } catch (FileNotFoundException e) {
            return "FileNotFoundException: " + e.getMessage();
        } catch (IOException e) {
            return "IOException: " + e.getMessage();
        } catch (SQLException e) {
            return "SQLException: " + e.getMessage();
        }
    }

    public static void main(String[] args) {
        //Konfiguriere und öffne MainWindow
        JFrame frame = new JFrame("MainWindow");
        frame.setContentPane(new MainWindow().panel_Main);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1000, 800);
        frame.setVisible(true);
        frame.setTitle("Protokollmanager Advanced");
        //TODO: Close database connection when window is closed
        /*frame.addWindowListener(new java.awt.event.WindowAdapter()
        {
            public void windowClosing(WindowEvent winEvt)
            {
                if (connection != null)
                    connection.close();
                System.exit(0);
            }
        });*/
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
