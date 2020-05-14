package com.gui;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
    private JButton btn_Script; //RT  2.43a
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
        setStyle();

        //Connecte zur Datenbank
        String str_connect = connectDatabase(str_DBPath);
        if (str_connect.equals("verbunden"))
            txtFieldStatus.setText(" Erfolgreich verbunden.");
        else
            txtFieldStatus.setText(str_connect);

        cmb_Query.addItem("CheckAllHazardClass.sql");
        cmb_Query.addItem("CheckAllStandardTestDatum.sql");
        cmb_Query.addItem("CheckAllSafetyTestDatum.sql");
        cmb_Query.addItem("CheckAllPrüfberichtDatum.sql");
        cmb_Query.setSelectedIndex(0);

        table_Response.setAutoCreateRowSorter(true);
        table_Response.setGridColor(Color.white);

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
            System.out.println(sheet.getRow(lastRow).getCell(10));

            workbook.close();
            txtFieldStatus.setText(" " + selectedFile.getName() + " wurde erfolgreich importiert");
        } catch (FileNotFoundException ex) {
            txtFieldStatus.setText(" FileNotFoundException: " + ex.getMessage());
        } catch (IOException ex) {
            txtFieldStatus.setText(" IOException: " + ex.getMessage());
        }
    }

    public void setStyle()
    {
        Color color_Background = new Color(32,136,203);
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
    }

    public String executeScript(String str_ScriptBasePath)
    {
        try {
            /*ScriptRunner sr = new ScriptRunner(connection); //using mybatis
            Reader reader = null;
            reader = new BufferedReader(new FileReader(str_ScriptBasePath + cmb_Query.getSelectedItem().toString()));
            sr.setSendFullScript(true);
            sr.runScript(reader);
            reader.close();*/
            //TODO: Return SQL RESPONSE https://stackoverflow.com/questions/8708342/redirect-console-output-to-string-in-java
            Path path = Paths.get(str_ScriptBasePath + cmb_Query.getSelectedItem().toString());
            String content = Files.readString(path, StandardCharsets.ISO_8859_1);
            PreparedStatement preparedStatement = connection.prepareStatement(content);

            try (ResultSet resultSet = preparedStatement.executeQuery())
            {
                ResultSetMetaData rsmd = resultSet.getMetaData();
                int columnCount = rsmd.getColumnCount();
                DefaultTableModel dtm = new DefaultTableModel(){
                    @Override
                    public boolean isCellEditable(int row, int column) {
                        return false;
                    }
                };

                for (int i = 1; i <= columnCount; i++)
                {
                    dtm.addColumn(rsmd.getColumnName(i));
                }
                while(resultSet.next())
                {
                    Vector<Object> data = new Vector<Object>();
                    for (int i = 1; i <= columnCount; i++)
                    {
                        data.add(resultSet.getString(i));

                    }
                    dtm.addRow(data);
                }
                table_Response.setModel(dtm);
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
