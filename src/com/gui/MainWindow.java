package com.gui;

import org.apache.ibatis.jdbc.ScriptRunner;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.sql.*;

public class MainWindow {
    private JButton btn_Send;
    private JTextArea txtArea_Result;
    private JComboBox cmb_Query;
    private JPanel panel_Main;
    private JComboBox comboBox1;

    public Connection connection;

    public MainWindow() {
        String str_DBPath = "jdbc:firebirdsql://localhost:3050/C:/Users/Public/Documents/MEBEDO/PROTOKOLLmanager8/DB/BackupEdit/Datenbank.FDB";
        String str_ScriptBasePath = "C:/Users/Public/Documents/Protokollmanger Advanced/";

        //Connecte zur Datenbank
        String str_connect = connectDatabase(str_DBPath);
        if (str_connect.equals("verbunden"))
            txtArea_Result.setText("Erfolgreich verbunden.");
        else
            txtArea_Result.setText(str_connect);

        cmb_Query.addItem("CheckAllHazardClass.sql");
        cmb_Query.addItem("CheckAllStandardTestDatum.sql");
        cmb_Query.addItem("CheckAllSafetyTestDatum.sql");
        cmb_Query.addItem("CheckAllPrüfberichtDatum.sql");
        cmb_Query.setSelectedIndex(0);

        btn_Send.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String str_result = executeScript(str_ScriptBasePath);
                txtArea_Result.setText(str_result);
            }
        });
    }

    public String executeScript(String str_ScriptBasePath)
    {
        try {
            ScriptRunner sr = new ScriptRunner(connection); //using mybatis
            Reader reader = null;
            reader = new BufferedReader(new FileReader(str_ScriptBasePath + cmb_Query.getSelectedItem().toString()));
            sr.setSendFullScript(true);
            sr.runScript(reader);
            reader.close();
            //TODO: Return SQL RESPONSE https://stackoverflow.com/questions/8708342/redirect-console-output-to-string-in-java
            return null;
        } catch (FileNotFoundException e) {
            return "FileNotFoundException: " + e.getMessage();
        } catch (IOException e)
        {
            return "IOException: " + e.getMessage();
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
            return "Verbindung erfolgreich hergestellt.";
        } catch (SQLException ex) {
            txtArea_Result.setText("SQLException: " + ex.getMessage());
            return "SQLException: " + ex.getMessage();
        }
    }
}
