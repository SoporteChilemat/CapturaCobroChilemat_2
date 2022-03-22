package Principal;

import Clases.Cruze;
import Clases.Estadistica;
import Clases.Totales;
import Clases.AñoMesSaldo;
import DAO.DocumentoCobranzaDAO;
import DAO.TotalesDAO;
import DAO.consultaVersionDAO;
import static Principal.Logica.arrCruzeMuyMALA;
import static Principal.Logica.arrOC;
import static Principal.Logica.arrCruzeBAD;
import static Principal.Logica.arrCruzeOK;
import static Principal.Logica.arrSinCruzar;
import static Principal.Logica.leerExcel;
import static Principal.Logica.manejo;
import static Principal.NewJFrame.arrValoresRestaSuma0;
import static Principal.NewJFrame.arrValoresRestaSuma1;
import static Principal.NewJFrame.arrValoresRestaSuma2;
import static Principal.NewJFrame.arrValoresRestaSuma3;
import static Principal.NewJFrame.arrValoresRestaSuma4;
import static Principal.NewJFrame.canFindAllFridaysInRange;
import static Principal.NewJFrame.cargarTablas;
import static Principal.NewJFrame.filterheader;
import static Principal.NewJFrame.filterheader1;
import static Principal.NewJFrame.filterheader2;
import static Principal.NewJFrame.filterheader3;
import static Principal.NewJFrame.filterheader4;
import static Principal.NewJFrame.jButton2;
import static Principal.NewJFrame.jComboBox1;
import static Principal.NewJFrame.jLabel10;
import static Principal.NewJFrame.jLabel12;
import static Principal.NewJFrame.jLabel2;
import static Principal.NewJFrame.jLabel4;
import static Principal.NewJFrame.jLabel6;
import static Principal.NewJFrame.jLabel8;
import static Principal.NewJFrame.jTabbedPane1;
import static Principal.NewJFrame.jTable1;
import static Principal.NewJFrame.jTable2;
import static Principal.NewJFrame.jTable3;
import static Principal.NewJFrame.jTable4;
import static Principal.NewJFrame.jTable5;
import static Principal.NewJFrame.suma;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;
import java.util.Set;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.AbstractCellEditor;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.RowFilter;
import javax.swing.RowFilter.ComparisonType;
import static javax.swing.SwingConstants.CENTER;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import net.coderazzi.filters.gui.AutoChoices;
import net.coderazzi.filters.gui.TableFilterHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.ini4j.Ini;
import org.ini4j.Wini;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.data.category.DefaultCategoryDataset;
import org.openide.util.Exceptions;

public class NewJFrame extends javax.swing.JFrame {

    public static TableFilterHeader filterheader;
    public static TableFilterHeader filterheader1;
    public static TableFilterHeader filterheader2;
    public static TableFilterHeader filterheader3;
    public static TableFilterHeader filterheader4;
    public static String filtro = "";
    public static ArrayList<String> arrValoresRestaSuma0 = new ArrayList<>();
    public static ArrayList<String> arrValoresRestaSuma1 = new ArrayList<>();
    public static ArrayList<String> arrValoresRestaSuma2 = new ArrayList<>();
    public static ArrayList<String> arrValoresRestaSuma3 = new ArrayList<>();
    public static ArrayList<String> arrValoresRestaSuma4 = new ArrayList<>();
    public static int sumaTiempos = 0;

    public NewJFrame() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenu2 = new javax.swing.JMenu();
        jPanel1 = new javax.swing.JPanel();
        jPanel19 = new javax.swing.JPanel();
        jLabel34 = new javax.swing.JLabel();
        jButton8 = new javax.swing.JButton();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTable2 = new javax.swing.JTable();
        jScrollPane3 = new javax.swing.JScrollPane();
        jTable3 = new javax.swing.JTable();
        jScrollPane5 = new javax.swing.JScrollPane();
        jTable5 = new javax.swing.JTable();
        jScrollPane4 = new javax.swing.JScrollPane();
        jTable4 = new javax.swing.JTable();
        jPanel5 = new javax.swing.JPanel();
        jComboBox3 = new javax.swing.JComboBox<>();
        jButton3 = new javax.swing.JButton();
        jPanel6 = new javax.swing.JPanel();
        jScrollPane6 = new javax.swing.JScrollPane();
        jTable6 = new javax.swing.JTable();
        jPanel2 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jPanel4 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox<>();
        jButton2 = new javax.swing.JButton();
        jComboBox2 = new javax.swing.JComboBox<>();
        jMenuBar2 = new javax.swing.JMenuBar();
        jMenu3 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenuItem2 = new javax.swing.JMenuItem();

        jMenu1.setText("File");
        jMenuBar1.add(jMenu1);

        jMenu2.setText("Edit");
        jMenuBar1.add(jMenu2);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        jPanel1.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jPanel19.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jLabel34.setText("Actualizacion Necesaria");

        jButton8.setText("Actualizar");
        jButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton8ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel19Layout = new javax.swing.GroupLayout(jPanel19);
        jPanel19.setLayout(jPanel19Layout);
        jPanel19Layout.setHorizontalGroup(
            jPanel19Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel19Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel19Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(jButton8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel34, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel19Layout.setVerticalGroup(
            jPanel19Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel19Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel34)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton8)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jTabbedPane1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mousePressed(java.awt.event.MouseEvent evt) {
                jTabbedPane1MousePressed(evt);
            }
        });

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        jTabbedPane1.addTab("OK", jScrollPane1);

        jTable2.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane2.setViewportView(jTable2);

        jTabbedPane1.addTab("REVISAR", jScrollPane2);

        jTable3.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane3.setViewportView(jTable3);

        jTabbedPane1.addTab("MALA", jScrollPane3);

        jTable5.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane5.setViewportView(jTable5);

        jTabbedPane1.addTab("ORDEN DE CARGA", jScrollPane5);

        jTable4.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane4.setViewportView(jTable4);

        jTabbedPane1.addTab("TOTALES", jScrollPane4);

        jComboBox3.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jButton3.setText("Buscar Infromacion");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 0, Short.MAX_VALUE)
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 211, Short.MAX_VALUE)
        );

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jComboBox3, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jButton3, javax.swing.GroupLayout.DEFAULT_SIZE, 154, Short.MAX_VALUE))
                        .addGap(0, 986, Short.MAX_VALUE))
                    .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        jTabbedPane1.addTab("Graficos", jPanel5);

        jTable6.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane6.setViewportView(jTable6);

        jTabbedPane1.addTab("Oculta", jScrollPane6);

        jPanel2.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jLabel1.setText("Total:");

        jLabel2.setText("jLabel2");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jLabel2)
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(12, 12, 12)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(jLabel2))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jLabel3.setText("Total OK:");

        jLabel4.setForeground(new java.awt.Color(255, 102, 102));
        jLabel4.setText("jLabel4");

        jLabel5.setText("Total Revisar:");

        jLabel6.setForeground(new java.awt.Color(255, 102, 102));
        jLabel6.setText("jLabel6");

        jLabel7.setText("Total Mala:");

        jLabel8.setForeground(new java.awt.Color(255, 102, 102));
        jLabel8.setText("jLabel8");

        jLabel9.setText("Suma:");

        jLabel10.setForeground(new java.awt.Color(255, 102, 102));
        jLabel10.setText("jLabel10");

        jLabel11.setText("OC:");

        jLabel12.setForeground(new java.awt.Color(255, 102, 102));
        jLabel12.setText("jLabel12");

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel4)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel5)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel6)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel7)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel8)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel11)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel12)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel9)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel10)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(12, 12, 12)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel9)
                    .addComponent(jLabel10)
                    .addComponent(jLabel8)
                    .addComponent(jLabel7)
                    .addComponent(jLabel6)
                    .addComponent(jLabel5)
                    .addComponent(jLabel4)
                    .addComponent(jLabel3)
                    .addComponent(jLabel11)
                    .addComponent(jLabel12))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel4.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jButton1.setText("Quitar Filtro");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setText("Filtrar");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "TODOS", "CALLE COLEGIO N° 1867 , OLMUE", "CALLE LOS QUILLAYES 1950, PEÑABLANCA", "MATURANA N° 90, VILLA ALEMANA", "PARCELA 65 CAMINO TRONCAL 1550 PEÑABLANCA", "BARRIO INDUSTRIAL 701 ALTO PEÑUELAS BODEGA E28" }));

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 80, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel19, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jTabbedPane1))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel19, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 311, Short.MAX_VALUE)
                .addContainerGap())
        );

        jMenu3.setText("Menu");

        jMenuItem1.setText("Exportar");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu3.add(jMenuItem1);

        jMenuItem2.setText("Cargar Archivo");
        jMenuItem2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem2ActionPerformed(evt);
            }
        });
        jMenu3.add(jMenuItem2);

        jMenuBar2.add(jMenu3);

        setJMenuBar(jMenuBar2);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    public static void compruebaVersion() {
        Thread thread = new Thread(() -> {
            while (opcion) {
                String data = "";
                Wini ini = null;
                FileInputStream inputStream = null;
                try {
                    inputStream = new FileInputStream(System.getProperty("user.dir") + "\\config.ini");
                    ini = new Wini(inputStream);
                    data = ini.get("Version", "v");
                    inputStream.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                }

                ArrayList<String> ConsultaVersion = null;
                try {
                    ConsultaVersion = consultaVersionDAO.ConsultaVersion();
                } catch (IOException ex) {
//                    Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                } catch (SQLException ex) {
//                    Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                }

                System.out.println(ConsultaVersion);

                for (int j = 0; j < ConsultaVersion.size(); j++) {
                    String get = ConsultaVersion.get(j);

                    if (Integer.valueOf(data) < Integer.valueOf(get)) {
                        v = get;
                        jPanel19.setVisible(true);
                    } else {
                        v = "";
                        jPanel19.setVisible(false);
                    }
                }

                try {
                    Thread.sleep(1000);
                } catch (InterruptedException ex) {
                    Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });
        thread.start();
    }

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        CompletableFuture.runAsync(() -> {
            try {
                compruebaVersion();
            } catch (Exception ex) {
            }

            jTabbedPane1.remove(6);
            ImageIcon imgThisImg = new ImageIcon(System.getProperty("user.dir") + "\\Iconos\\alert-icon.png");
            Image image = imgThisImg.getImage(); // transform it 
            Image newimg = image.getScaledInstance(20, 20, java.awt.Image.SCALE_SMOOTH); // scale it the smooth way  
            imgThisImg = new ImageIcon(newimg);  // transform it back
            jLabel34.setIcon(imgThisImg);

            try {
                jTabbedPane1.setVisible(false);
                jPanel2.setVisible(false);
                jPanel3.setVisible(false);
                jPanel4.setVisible(false);
                try {
                    int size = arrCruzeOK.size();
                    System.out.println(size);
                    int size1 = arrCruzeBAD.size();
                    System.out.println(size1);
                    int size2 = arrCruzeMuyMALA.size();
                    System.out.println(size2);

                    filterheader = new TableFilterHeader(jTable1, AutoChoices.ENABLED);
                    filterheader1 = new TableFilterHeader(jTable2, AutoChoices.ENABLED);
                    filterheader2 = new TableFilterHeader(jTable3, AutoChoices.ENABLED);
                    filterheader3 = new TableFilterHeader(jTable4, AutoChoices.ENABLED);
                    filterheader4 = new TableFilterHeader(jTable5, AutoChoices.ENABLED);

                    cargarTablas();

                    canFindAllFridaysInRange(jTable1, 1);

                } catch (ParseException ex) {
                    Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                }
                jLabel2.setText("" + jTable1.getRowCount());

                llenaTabla4();

                suma();
                selectedTabla4();
                corrigueSelected();
                todosClientes();

                AutoCompletion.enable(jComboBox3);
            } catch (ParseException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            } catch (SQLException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
        }).thenRun(() -> {
            jTabbedPane1.setVisible(true);
            jPanel2.setVisible(true);
            jPanel3.setVisible(true);
            jPanel4.setVisible(true);
        });
    }//GEN-LAST:event_formWindowOpened

    public static void todosClientes() {
        int rowCount = jTable1.getRowCount();
        int rowCount1 = jTable2.getRowCount();
        int rowCount2 = jTable3.getRowCount();
        int rowCount3 = jTable5.getRowCount();

        ArrayList<String> arrProveedores = new ArrayList<>();

        for (int i = 0; i < rowCount; i++) {
            String toString = jTable1.getValueAt(i, 7).toString();
            arrProveedores.add(toString.trim());
        }

        for (int i = 0; i < rowCount1; i++) {
            String toString = jTable2.getValueAt(i, 7).toString();
            arrProveedores.add(toString.trim());
        }

        for (int i = 0; i < rowCount2; i++) {
            String toString = jTable3.getValueAt(i, 7).toString();
            arrProveedores.add(toString.trim());
        }

        for (int i = 0; i < rowCount3; i++) {
            String toString = jTable5.getValueAt(i, 7).toString();
            arrProveedores.add(toString.trim());
        }

        Set<String> hashSet = new HashSet<String>(arrProveedores);
        arrProveedores.clear();
        arrProveedores.addAll(hashSet);

        System.out.println(arrProveedores);
        jComboBox3.removeAllItems();

        Collections.sort(arrProveedores);

        jComboBox3.addItem("TODOS");
        arrProveedores.stream().forEach((String nombre) -> {
            jComboBox3.addItem(nombre);
        });
    }

    public static void llenaTabla4() throws ParseException, IOException, SQLException {
        ArrayList<ArrayList<String>> canFindAllFridaysInRange = NewJFrame.canFindAllFridaysInRange(jTable1, 0);
        ArrayList<ArrayList<String>> canFindAllFridaysInRange1 = NewJFrame.canFindAllFridaysInRange(jTable2, 0);
        ArrayList<ArrayList<String>> canFindAllFridaysInRange2 = NewJFrame.canFindAllFridaysInRange(jTable3, 0);
        ArrayList<ArrayList<String>> canFindAllFridaysInRange3 = NewJFrame.canFindAllFridaysInRange(jTable5, 0);

        for (int i = 0; i < canFindAllFridaysInRange.size(); i++) {
            ArrayList<String> get = canFindAllFridaysInRange.get(i);
            String get1 = get.get(0);
            String get2 = get.get(1);

            for (int j = 0; j < canFindAllFridaysInRange1.size(); j++) {
                ArrayList<String> get3 = canFindAllFridaysInRange1.get(j);
                String get11 = get3.get(0);
                String get22 = get3.get(1);

                if (get1.equals(get11) && get2.equals(get22)) {
                    canFindAllFridaysInRange1.remove(j);
                }
            }
        }

        for (int i = 0; i < canFindAllFridaysInRange.size(); i++) {
            ArrayList<String> get = canFindAllFridaysInRange.get(i);
            String get1 = get.get(0);
            String get2 = get.get(1);

            for (int j = 0; j < canFindAllFridaysInRange2.size(); j++) {
                ArrayList<String> get3 = canFindAllFridaysInRange2.get(j);
                String get11 = get3.get(0);
                String get22 = get3.get(1);

                if (get1.equals(get11) && get2.equals(get22)) {
                    canFindAllFridaysInRange2.remove(j);
                }
            }
        }

        for (int i = 0; i < canFindAllFridaysInRange.size(); i++) {
            ArrayList<String> get = canFindAllFridaysInRange.get(i);
            String get1 = get.get(0);
            String get2 = get.get(1);

            for (int j = 0; j < canFindAllFridaysInRange3.size(); j++) {
                ArrayList<String> get3 = canFindAllFridaysInRange3.get(j);
                String get11 = get3.get(0);
                String get22 = get3.get(1);

                if (get1.equals(get11) && get2.equals(get22)) {
                    canFindAllFridaysInRange3.remove(j);
                }
            }
        }

        System.out.println("canFindAllFridaysInRange1 " + canFindAllFridaysInRange1);

        if (!canFindAllFridaysInRange1.isEmpty()) {
            for (int i = 0; i < canFindAllFridaysInRange1.size(); i++) {
                boolean bool = false;

                ArrayList<String> get = canFindAllFridaysInRange1.get(i);
                String get1 = get.get(1);
                String[] split = get1.split("/");
                Integer valueOf = Integer.valueOf(split[2] + "" + split[1] + "" + split[0]);

                for (int j = 0; j < canFindAllFridaysInRange.size(); j++) {
                    ArrayList<String> get2 = canFindAllFridaysInRange.get(j);
                    String get3 = get2.get(1);
                    String[] split1 = get3.split("/");

                    Integer valueOf1 = Integer.valueOf(split1[2] + "" + split1[1] + "" + split1[0]);

                    if (valueOf1 > valueOf) {
                        System.out.println("valueOf " + valueOf + " valueOf1 " + valueOf1);
                        canFindAllFridaysInRange.add(j, get);
                        bool = true;
                        break;
                    }
                }
                if (bool == false) {
                    canFindAllFridaysInRange.add(get);
                }
                bool = false;
            }
        }

        System.out.println("canFindAllFridaysInRange2 " + canFindAllFridaysInRange2);

        if (!canFindAllFridaysInRange2.isEmpty()) {
            for (int i = 0; i < canFindAllFridaysInRange2.size(); i++) {
                boolean bool = false;

                ArrayList<String> get = canFindAllFridaysInRange2.get(i);
                String get1 = get.get(1);
                String[] split = get1.split("/");
                Integer valueOf = Integer.valueOf(split[2] + "" + split[1] + "" + split[0]);

                for (int j = 0; j < canFindAllFridaysInRange.size(); j++) {
                    ArrayList<String> get2 = canFindAllFridaysInRange.get(j);
                    String get3 = get2.get(1);
                    String[] split1 = get3.split("/");

                    Integer valueOf1 = Integer.valueOf(split1[2] + "" + split1[1] + "" + split1[0]);

                    if (valueOf1 > valueOf) {
                        System.out.println("valueOf " + valueOf + " valueOf1 " + valueOf1);
                        canFindAllFridaysInRange.add(j, get);
                        bool = true;
                        break;
                    }
                }

                if (bool == false) {
                    canFindAllFridaysInRange.add(get);
                }
                bool = false;
            }
        }

        System.out.println("canFindAllFridaysInRange2 " + canFindAllFridaysInRange2);

        if (!canFindAllFridaysInRange3.isEmpty()) {
            for (int i = 0; i < canFindAllFridaysInRange3.size(); i++) {
                boolean bool = false;

                ArrayList<String> get = canFindAllFridaysInRange3.get(i);
                String get1 = get.get(1);
                String[] split = get1.split("/");
                Integer valueOf = Integer.valueOf(split[2] + "" + split[1] + "" + split[0]);

                for (int j = 0; j < canFindAllFridaysInRange.size(); j++) {
                    ArrayList<String> get2 = canFindAllFridaysInRange.get(j);
                    String get3 = get2.get(1);
                    String[] split1 = get3.split("/");

                    Integer valueOf1 = Integer.valueOf(split1[2] + "" + split1[1] + "" + split1[0]);

                    if (valueOf1 > valueOf) {
                        System.out.println("valueOf " + valueOf + " valueOf1 " + valueOf1);
                        canFindAllFridaysInRange.add(j, get);
                        bool = true;
                        break;
                    }
                }

                if (bool == false) {
                    canFindAllFridaysInRange.add(get);
                }
                bool = false;
            }
        }
        System.out.println("canFindAllFridaysInRange3 " + canFindAllFridaysInRange3);

        System.out.println(canFindAllFridaysInRange);
        DefaultTableModel model = new DefaultTableModel();

        model.addColumn("Fecha");
        model.addColumn("Total OK");
        model.addColumn("Total Revisar");
        model.addColumn("Total Mala");
        model.addColumn("OrdenDeCarga");
        model.addColumn("Suma");
        model.addColumn("Pagado/No Pagado");
        model.addColumn("Comentario");

        Object[] fila = new Object[8];

        int size1 = canFindAllFridaysInRange.size();
        System.out.println("size1 " + size1);

        for (int i = 0; i < canFindAllFridaysInRange.size(); i++) {
            ArrayList<String> get = canFindAllFridaysInRange.get(i);
            String get1 = get.get(0);
            System.out.println("get1 " + get1);

            Date date1 = new SimpleDateFormat("dd/MM/yyyy").parse(get1);
            String get2 = get.get(1);
            System.out.println("get2 " + get2);

            Date date2 = new SimpleDateFormat("dd/MM/yyyy").parse(get2);

            filter(date1, date2, jTable1);
            filter(date1, date2, jTable2);
            filter(date1, date2, jTable3);
            filter(date1, date2, jTable5);

            String[] valores = suma();
            System.out.println("suma " + valores[3]);

            fila[0] = "Sabado: " + get1 + " Viernes: " + get2;
            fila[1] = valores[0];
            fila[2] = valores[1];
            fila[3] = valores[2];
            fila[4] = valores[3];
            fila[5] = valores[4];

            ArrayList<Totales> consultaTotales = TotalesDAO.consultaTotales("ingresos");
            System.out.println("consultaTotales.size() " + consultaTotales.size());

            boolean bool = false;
            int num = 999999999;
            String comentario = "";

            for (int j = 0; j < consultaTotales.size(); j++) {
                Totales total = consultaTotales.get(j);
                String fechas = total.getFechas();
                int selected = total.getSelected();
                String comentario1 = total.getComentario();

                System.out.println("fechas " + fechas);
                System.out.println("fechas " + "Sabado: " + get1 + " Viernes: " + get2);

                if (fechas.equals("Sabado: " + get1 + " Viernes: " + get2)) {
                    num = selected;
                    bool = true;
                    if (comentario1 != null) {
                        comentario = comentario1;
                    } else {
                        comentario = "";
                    }
                    break;
                }
            }

            if (bool == true) {
                if (num == 0) {
                    fila[6] = true;
                } else if (num == 1) {
                    fila[6] = false;
                }
            } else {
                fila[6] = false;
            }

            bool = false;
            num = 999999999;

            fila[7] = comentario;

            model.addRow(fila);
        }

        jTable4.setModel(model);

        DefaultTableCellRenderer Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

        jTable4.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable4.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable4.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable4.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable4.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable4.getColumnModel().getColumn(5).setCellRenderer(Renderer);

        jTable4.getColumnModel().getColumn(1).setPreferredWidth(5);
        jTable4.getColumnModel().getColumn(2).setPreferredWidth(5);
        jTable4.getColumnModel().getColumn(3).setPreferredWidth(5);
        jTable4.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable4.getColumnModel().getColumn(5).setPreferredWidth(5);

        jTable4.getColumn("Pagado/No Pagado").setCellRenderer(new Render_CheckBox());
        jTable4.getColumn("Pagado/No Pagado").setCellEditor(new Editor_CheckBox());

        jTable4.getColumn("Comentario").setCellRenderer(new CustomRenderer(jTable4));
        jTable4.getColumn("Comentario").setCellEditor(new CustomEditor(jTable4));

        jTable4.setRowHeight(40);

        jTable1.setRowSorter(null);
        jTable2.setRowSorter(null);
        jTable3.setRowSorter(null);
        jTable5.setRowSorter(null);

        suma();

        TableModel modelx1 = jTable1.getModel();
        TableRowSorter trs = new TableRowSorter(modelx1);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable1.setRowSorter(trs);

        TableModel modelx2 = jTable2.getModel();
        trs = new TableRowSorter(modelx2);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable2.setRowSorter(trs);

        TableModel modelx3 = jTable3.getModel();
        trs = new TableRowSorter(modelx3);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable3.setRowSorter(trs);

        TableModel modelx4 = jTable5.getModel();
        trs = new TableRowSorter(modelx4);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable5.setRowSorter(trs);
    }

    public static String[] suma() {
        double suma = 0;
        double sumaTotal = 0;
        String[] arrValores = new String[5];

        for (int i = 0; i < jTable1.getRowCount(); i++) {
            String valor = jTable1.getValueAt(i, 10).toString();
            String resta = jTable1.getValueAt(i, 11).toString();
            valor = valor.replace(".", "").trim();
//            System.out.println("valor " + valor);
            double numero = Double.parseDouble(valor);
            double numeroResta = Double.parseDouble(resta);
            suma = numero + suma - numeroResta;
        }

        BigDecimal bd = new BigDecimal(suma);
//        System.out.println("suma " + bd);
        sumaTotal = suma + sumaTotal;

        suma = 0;

        DecimalFormat formato = new DecimalFormat("#,###");
        String valorFormateado = formato.format(bd);
//        System.out.println(valorFormateado);

        String valor1 = valorFormateado;

        jLabel4.setText("" + valorFormateado);
////////////////////////////////////////////////////////////////////////////////
        for (int i = 0; i < jTable2.getRowCount(); i++) {
            String valor = jTable2.getValueAt(i, 10).toString();
            String resta = jTable2.getValueAt(i, 11).toString();
            valor = valor.replace(".", "").trim();
//            System.out.println("valor " + valor);
            double numero = Double.parseDouble(valor);
            double numeroResta = Double.parseDouble(resta);
            suma = numero + suma - numeroResta;
        }

        bd = new BigDecimal(suma);
//        System.out.println("suma " + bd);
        sumaTotal = suma + sumaTotal;

        suma = 0;

        formato = new DecimalFormat("#,###");
        valorFormateado = formato.format(bd);
//        System.out.println(valorFormateado);

        String valor2 = valorFormateado;

        jLabel6.setText("" + valorFormateado);
////////////////////////////////////////////////////////////////////////////////
        for (int i = 0; i < jTable3.getRowCount(); i++) {
            String valor = jTable3.getValueAt(i, 10).toString();
            String resta = jTable3.getValueAt(i, 11).toString();
            valor = valor.replace(".", "").trim();
//            System.out.println("valor " + valor);
            double numero = Double.parseDouble(valor);
            double numeroResta = Double.parseDouble(resta);
            suma = numero + suma - numeroResta;
        }

        bd = new BigDecimal(suma);
//        System.out.println("suma " + bd);
        sumaTotal = suma + sumaTotal;

        suma = 0;

        formato = new DecimalFormat("#,###");
        valorFormateado = formato.format(bd);
//        System.out.println(valorFormateado);

        String valor3 = valorFormateado;

        jLabel8.setText("" + valorFormateado);
////////////////////////////////////////////////////////////////////////////////
        for (int i = 0; i < jTable5.getRowCount(); i++) {
            String valor = jTable5.getValueAt(i, 10).toString();
            String resta = jTable5.getValueAt(i, 11).toString();
            valor = valor.replace(".", "").trim();
//            System.out.println("valor " + valor);
            double numero = Double.parseDouble(valor);
            double numeroResta = Double.parseDouble(resta);
            suma = numero + suma - numeroResta;
        }

        bd = new BigDecimal(suma);
//        System.out.println("suma " + bd);
        sumaTotal = suma + sumaTotal;

        formato = new DecimalFormat("#,###");
        valorFormateado = formato.format(bd);
//        System.out.println(valorFormateado);

        String valor5 = valorFormateado;

        jLabel12.setText("" + valorFormateado);

        BigDecimal bdT = new BigDecimal(sumaTotal);

        formato = new DecimalFormat("#,###");
        valorFormateado = formato.format(bdT);
//        System.out.println(valorFormateado);

        String valor4 = valorFormateado;

        jLabel10.setText("" + valorFormateado);

        arrValores[0] = valor1;
        arrValores[1] = valor2;
        arrValores[2] = valor3;
        arrValores[3] = valor5;
        arrValores[4] = valor4;
        return arrValores;
    }

    public static ArrayList<ArrayList<String>> canFindAllFridaysInRange(JTable jTable, int o) throws ParseException {
        LocalDate start1 = LocalDate.of(2021, 1, 1);
        LocalDate end1 = LocalDate.of(2022, 12, 31);

        DayOfWeek dowOfStart1 = start1.getDayOfWeek();
        int difference1 = DayOfWeek.SATURDAY.getValue() - dowOfStart1.getValue();
        if (difference1 < 0) {
            difference1 += 7;
        }

        List<LocalDate> fridaysInRange1 = new ArrayList<>();
        List<LocalDate> fridaysInRange2 = new ArrayList<>();

        LocalDate currentFriday1 = start1.plusDays(difference1);

        do {
            fridaysInRange1.add(currentFriday1);
            LocalDate plusDays = currentFriday1.plusDays(6);

            currentFriday1 = currentFriday1.plusDays(7);
            fridaysInRange2.add(plusDays);
        } while (currentFriday1.isBefore(end1));

        ArrayList<ArrayList<String>> arrFechas = new ArrayList<>();

        AtomicInteger num = new AtomicInteger(0);
        fridaysInRange1.forEach((LocalDate get) -> {
            String toString = get.toString();
            String toString1 = fridaysInRange2.get(Integer.valueOf(num.toString())).toString();

            String[] split = toString.split("-");
            String[] split1 = toString1.split("-");

            String name = split[0] + "" + split[1] + "" + split[2];
            Integer sabado = Integer.valueOf(name);

            String name1 = split1[0] + "" + split1[1] + "" + split1[2];
            Integer viernes = Integer.valueOf(name1);

            boolean bool = false;

            for (int t = 0; t < jTable.getModel().getRowCount(); t++) {
                sumaTiempos = sumaTiempos + jTable.getModel().getRowCount();
                try {
                    String toString0 = jTable.getModel().getValueAt(t, 9).toString();

                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                    Date parse = dateFormat.parse(toString0);

                    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
                    String strDate = formatter.format(parse);

                    String[] split0 = strDate.split("/");
                    Integer valueOf2 = Integer.valueOf(split0[2] + "" + split0[1] + "" + split0[0]);

                    if (valueOf2 >= sabado && valueOf2 <= viernes) {
                        bool = true;
                        break;
                    }
                } catch (Exception ex) {
                }
            }

            if (bool) {
                ArrayList<String> arrParDeFechas = new ArrayList<>();
                if (o != 0) {
                    jComboBox1.addItem("Sabado: " + split[2] + "/" + split[1] + "/" + split[0] + " Viernes: " + split1[2] + "/" + split1[1] + "/" + split1[0]);
                } else {
                    arrParDeFechas.add(split[2] + "/" + split[1] + "/" + split[0]);
                    arrParDeFechas.add(split1[2] + "/" + split1[1] + "/" + split1[0]);
                    arrFechas.add(arrParDeFechas);
                }
            }
            bool = false;
            num.getAndIncrement();
        });
        return arrFechas;
    }

    public static void total() {
        int selectedIndex = jTabbedPane1.getSelectedIndex();

        switch (selectedIndex) {
            case 0:
                try {
                jComboBox1.removeAllItems();
                canFindAllFridaysInRange(jTable1, 1);
                jLabel2.setText("" + jTable1.getRowCount());
            } catch (ParseException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
            break;
            case 1:
                try {
                jComboBox1.removeAllItems();
                canFindAllFridaysInRange(jTable2, 1);
                jLabel2.setText("" + jTable2.getRowCount());
            } catch (ParseException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
            break;
            case 2:
                try {
                jComboBox1.removeAllItems();
                canFindAllFridaysInRange(jTable3, 1);
                jLabel2.setText("" + jTable3.getRowCount());
            } catch (ParseException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
            break;
            case 3:
                try {
                jComboBox1.removeAllItems();
                canFindAllFridaysInRange(jTable5, 1);
                jLabel2.setText("" + jTable5.getRowCount());
            } catch (ParseException ex) {
                Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
            }
            break;
            default:
                break;
        }
    }

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:        
        jButton2.setText("Filtrar");
        int selectedIndex = jTabbedPane1.getSelectedIndex();

        jTable1.setRowSorter(null);
        jTable2.setRowSorter(null);
        jTable3.setRowSorter(null);
        jTable5.setRowSorter(null);

        if (selectedIndex == 0) {
            jLabel2.setText("" + jTable1.getRowCount());
        } else if (selectedIndex == 1) {
            jLabel2.setText("" + jTable2.getRowCount());
        } else if (selectedIndex == 2) {
            jLabel2.setText("" + jTable3.getRowCount());
        } else if (selectedIndex == 3) {
            jLabel2.setText("" + jTable5.getRowCount());
        }

        suma();
        selectedTabla4();
        corrigueSelected();

        filtro = "Filtrar";

        TableModel modelx1 = jTable1.getModel();
        TableRowSorter trs = new TableRowSorter(modelx1);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable1.setRowSorter(trs);

        TableModel modelx2 = jTable2.getModel();
        trs = new TableRowSorter(modelx2);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable2.setRowSorter(trs);

        TableModel modelx3 = jTable3.getModel();
        trs = new TableRowSorter(modelx3);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable3.setRowSorter(trs);

        TableModel modelx4 = jTable5.getModel();
        trs = new TableRowSorter(modelx4);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

        jTable5.setRowSorter(trs);
    }//GEN-LAST:event_jButton1ActionPerformed

    public void filtrarLocal() {
        try {
            int selectedIndex = jTabbedPane1.getSelectedIndex();

//            if (selectedIndex == 0) {
            String toString = jComboBox1.getSelectedItem().toString();
            System.out.println("" + toString);

            String substring = toString.substring(8, 18).trim();
            System.out.println("substring " + substring);
            Date date1 = new SimpleDateFormat("dd/MM/yyyy").parse(substring);

            String substring1 = toString.substring(28, 38).trim();
            System.out.println("substring1 " + substring1);
            Date date11 = new SimpleDateFormat("dd/MM/yyyy").parse(substring1);

            filter(date1, date11, jTable1);
            filter(date1, date11, jTable2);
            filter(date1, date11, jTable3);

            if (selectedIndex == 0) {
                jLabel2.setText("" + jTable1.getRowCount());
            }
            if (selectedIndex == 1) {
                jLabel2.setText("" + jTable2.getRowCount());
            }
            if (selectedIndex == 2) {
                jLabel2.setText("" + jTable3.getRowCount());
            }
        } catch (ParseException ex) {
            Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static void filtrarSemana() {
        try {
            int selectedIndex = jTabbedPane1.getSelectedIndex();
            String toString = jComboBox1.getSelectedItem().toString();
            System.out.println("" + toString);

            String substring = toString.substring(8, 18).trim();
            System.out.println("substring " + substring);
            Date date1 = new SimpleDateFormat("dd/MM/yyyy").parse(substring);

            String substring1 = toString.substring(28, 38).trim();
            System.out.println("substring1 " + substring1);
            Date date11 = new SimpleDateFormat("dd/MM/yyyy").parse(substring1);

            filter(date1, date11, jTable1);
            filter(date1, date11, jTable2);
            filter(date1, date11, jTable3);
            filter(date1, date11, jTable5);

            if (selectedIndex == 0) {
                jLabel2.setText("" + jTable1.getRowCount());
            }
            if (selectedIndex == 1) {
                jLabel2.setText("" + jTable2.getRowCount());
            }
            if (selectedIndex == 2) {
                jLabel2.setText("" + jTable3.getRowCount());
            }
            if (selectedIndex == 3) {
                jLabel2.setText("" + jTable5.getRowCount());
            }
        } catch (ParseException ex) {
            Exceptions.printStackTrace(ex);
        }
    }

    public static void filtrarSemana2() {
        try {
            if (!filtro.equals("Filtrar")) {
                String toString = filtro;
                int selectedIndex = jTabbedPane1.getSelectedIndex();
                String substring = toString.substring(8, 18).trim();
                System.out.println("substring " + substring);
                Date date1 = new SimpleDateFormat("dd/MM/yyyy").parse(substring);

                String substring1 = toString.substring(28, 38).trim();
                System.out.println("substring1 " + substring1);
                Date date11 = new SimpleDateFormat("dd/MM/yyyy").parse(substring1);

                filter(date1, date11, jTable1);
                filter(date1, date11, jTable2);
                filter(date1, date11, jTable3);
                filter(date1, date11, jTable5);

                if (selectedIndex == 0) {
                    jLabel2.setText("" + jTable1.getRowCount());
                }
                if (selectedIndex == 1) {
                    jLabel2.setText("" + jTable2.getRowCount());
                }
                if (selectedIndex == 2) {
                    jLabel2.setText("" + jTable3.getRowCount());
                }
                if (selectedIndex == 3) {
                    jLabel2.setText("" + jTable5.getRowCount());
                }
            }
        } catch (ParseException ex) {
            Exceptions.printStackTrace(ex);
        }
    }

    private void jTabbedPane1MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jTabbedPane1MousePressed
        int selectedIndex = jTabbedPane1.getSelectedIndex();
        CompletableFuture.runAsync(() -> {
            try {
                jTabbedPane1.setEnabled(false);
                jTable1.setVisible(false);
                String property = System.getProperty("user.dir");
                File file = new File(property + "\\Iconos\\loading_2.gif");
                Image image = Toolkit.getDefaultToolkit().createImage(org.apache.commons.io.IOUtils.toByteArray(new FileInputStream(file)));
                ImageIcon icon = new ImageIcon(image);
                jTabbedPane1.setIconAt(selectedIndex, icon);
                jTabbedPane1.getSelectedIndex();
                total();
                suma();
                selectedTabla4();
                corrigueSelected();
            } catch (IOException ex) {
                Exceptions.printStackTrace(ex);
            }
        }).thenRun(() -> {
            try {
                jTabbedPane1.setEnabled(true);
                jTable1.setVisible(true);
                String property = System.getProperty("user.dir");
                File file = new File(property + "\\Iconos\\vacio.png");
                Image image = Toolkit.getDefaultToolkit().createImage(org.apache.commons.io.IOUtils.toByteArray(new FileInputStream(file)));
                ImageIcon icon = new ImageIcon(image);
                jTabbedPane1.setIconAt(selectedIndex, icon);
            } catch (FileNotFoundException ex) {
                Exceptions.printStackTrace(ex);
            } catch (IOException ex) {
                Exceptions.printStackTrace(ex);
            }
        });
    }//GEN-LAST:event_jTabbedPane1MousePressed

    public static void selectedTabla4() {
        arrValoresRestaSuma0.clear();
        arrValoresRestaSuma1.clear();
        arrValoresRestaSuma2.clear();
        arrValoresRestaSuma3.clear();
        arrValoresRestaSuma4.clear();

        for (int i = 0; i < jTable4.getRowCount(); i++) {
            String toString = jTable4.getValueAt(i, 6).toString();
            if (toString.equals("true")) {
                String toString1 = jTable4.getValueAt(i, 1).toString();
                String toString2 = jTable4.getValueAt(i, 2).toString();
                String toString3 = jTable4.getValueAt(i, 3).toString();
                String toString4 = jTable4.getValueAt(i, 4).toString();
                String toString5 = jTable4.getValueAt(i, 5).toString();

                arrValoresRestaSuma0.add(toString1);
                arrValoresRestaSuma1.add(toString2);
                arrValoresRestaSuma2.add(toString3);
                arrValoresRestaSuma3.add(toString4);
                arrValoresRestaSuma4.add(toString5);
            }
        }
    }

    public static void corrigueSelected() {
        if (jButton2.getText().equals("Filtrar")) {
            String text = jLabel10.getText();
            double valueOf = Double.valueOf(text.replace(".", ""));

            double Suma = 0;
            for (int i = 0; i < arrValoresRestaSuma4.size(); i++) {
                String num = arrValoresRestaSuma4.get(i);
                double valueOf0 = Double.valueOf(num.replace(".", ""));
                Suma = Suma + valueOf0;
            }

            double name = valueOf - Suma;

            DecimalFormat formato = new DecimalFormat("#,###");
            String valorFormateado = formato.format(name);
            jLabel10.setText(valorFormateado);
            ////////////////////////////////////////////////////////////
            text = jLabel4.getText();
            valueOf = Double.valueOf(text.replace(".", ""));

            Suma = 0;
            for (int i = 0; i < arrValoresRestaSuma0.size(); i++) {
                String num = arrValoresRestaSuma0.get(i);
                double valueOf0 = Double.valueOf(num.replace(".", ""));
                Suma = Suma + valueOf0;
            }

            name = valueOf - Suma;

            formato = new DecimalFormat("#,###");
            valorFormateado = formato.format(name);
            jLabel4.setText(valorFormateado);
            //////////
            text = jLabel6.getText();
            valueOf = Double.valueOf(text.replace(".", ""));

            Suma = 0;
            for (int i = 0; i < arrValoresRestaSuma1.size(); i++) {
                String num = arrValoresRestaSuma1.get(i);
                double valueOf0 = Double.valueOf(num.replace(".", ""));
                Suma = Suma + valueOf0;
            }

            name = valueOf - Suma;

            formato = new DecimalFormat("#,###");
            valorFormateado = formato.format(name);
            jLabel6.setText(valorFormateado);
            ////////////
            text = jLabel8.getText();
            valueOf = Double.valueOf(text.replace(".", ""));

            Suma = 0;
            for (int i = 0; i < arrValoresRestaSuma2.size(); i++) {
                String num = arrValoresRestaSuma2.get(i);
                double valueOf0 = Double.valueOf(num.replace(".", ""));
                Suma = Suma + valueOf0;
            }

            name = valueOf - Suma;

            formato = new DecimalFormat("#,###");
            valorFormateado = formato.format(name);
            jLabel8.setText(valorFormateado);
            //////////////
            text = jLabel12.getText();
            valueOf = Double.valueOf(text.replace(".", ""));

            Suma = 0;
            for (int i = 0; i < arrValoresRestaSuma3.size(); i++) {
                String num = arrValoresRestaSuma3.get(i);
                double valueOf0 = Double.valueOf(num.replace(".", ""));
                Suma = Suma + valueOf0;
            }
            name = valueOf - Suma;

            formato = new DecimalFormat("#,###");
            valorFormateado = formato.format(name);
            jLabel12.setText(valorFormateado);
        }
    }


    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        // TODO add your handling code here:
        filtro = jComboBox1.getSelectedItem().toString();
        jButton2.setText(filtro);
        filtrarSemana();

//        DefaultTableModel modelFilter = new DefaultTableModel();
//
//        modelFilter.addColumn("pkNumeroCuota");
//        modelFilter.addColumn("numeroOrdenDoc");
//        modelFilter.addColumn("guiaChilemat");
//        modelFilter.addColumn("guiaProveedor");
//        modelFilter.addColumn("local");
//        modelFilter.addColumn("numeroOrdenIngreso");
////        model.addColumn("numeroGuia");
////        model.addColumn("fechaRecepcion");
//        modelFilter.addColumn("sucursal");
//        modelFilter.addColumn("proveedor");
//        modelFilter.addColumn("fechaEmision");
//        modelFilter.addColumn("fechaVencimiento");
////        model.addColumn("total");
//        modelFilter.addColumn("saldo");
//        modelFilter.addColumn("Nota de Credito");
//        modelFilter.addColumn("totalNCuota");
//        modelFilter.addColumn("procentaje");
//        modelFilter.addColumn("comentario");
//        modelFilter.addColumn("estado");
//
//        for (int i = 0; i < jTable1.getRowCount(); i++) {
//            Object[] fila = new Object[16];
//            for (int j = 0; j < jTable1.getColumnCount(); j++) {
//                Object valueAt = jTable1.getValueAt(i, j);
//                fila[j] = valueAt;
//            }
//            modelFilter.addRow(fila);
//        }
//        jTable1.removeAll();
//        jTable1.setModel(modelFilter);
//
//        TableRowSorter trs = new TableRowSorter(modelFilter);
//
//        trs.setComparator(8, new IntComparatorFecha0());
//        trs.setComparator(9, new IntComparatorFecha1());
//
//        trs.setComparator(10, new IntComparator());
//        trs.setComparator(11, new IntComparator());
//        trs.setComparator(12, new IntComparator());
//
////        trs.setComparator(12, new IntComparator());
////        trs.setComparator(2, new StringComparator2());
////        trs.setComparator(3, new IntComparator());
////        trs.setComparator(4, new StringComparator());
////        trs.setComparator(5, new StringComparator());
////        trs.setComparator(6, new StringComparator());
//        jTable1.setRowSorter(trs);
////                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
////                renderer.setHorizontalAlignment(JLabel.CENTER);
//
//        DefaultTableCellRenderer centerRenderer2 = new DefaultTableCellRenderer() {
//            @Override
//            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
//                try {
//                    DateFormat dateFormat0 = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
//                    Date parse = dateFormat0.parse(value.toString());
//                    DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
//                    String strDate = dateFormat.format(parse);
//                    value = strDate;
//                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
//                } catch (ParseException ex) {
//                    System.out.println(ex);
//                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
//                }
//            }
//        };
//        centerRenderer2.setHorizontalAlignment(JLabel.CENTER);
//
//        DefaultTableCellRenderer Renderer = new DefaultTableCellRenderer();
//        Renderer.setHorizontalAlignment(JLabel.CENTER);
//
////        jTable1.getColumnModel().getColumn(0).setWidth(0);
////        jTable1.getColumnModel().getColumn(0).setMinWidth(0);
////        jTable1.getColumnModel().getColumn(0).setMaxWidth(0);
//        jTable1.getColumnModel().getColumn(0).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(1).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(2).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(3).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(4).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(5).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(6).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(7).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(8).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(2).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
//        jTable1.getColumnModel().getColumn(10).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(11).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(12).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(13).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(14).setCellRenderer(Renderer);
//        jTable1.getColumnModel().getColumn(15).setCellRenderer(Renderer);
//
//        jTable1.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);
//
//        jTable1.getColumnModel().getColumn(1).setPreferredWidth(30);
//        jTable1.getColumnModel().getColumn(2).setPreferredWidth(30);
//        jTable1.getColumnModel().getColumn(3).setPreferredWidth(30);
//        jTable1.getColumnModel().getColumn(4).setPreferredWidth(5);
//        jTable1.getColumnModel().getColumn(11).setPreferredWidth(150);
//        jTable1.getColumnModel().getColumn(14).setPreferredWidth(250);
//        jTable1.getColumnModel().getColumn(15).setPreferredWidth(75);
//
//        jTable1.setRowHeight(80);
//        jTable1.setShowHorizontalLines(true);
//        jTable1.setShowVerticalLines(true);
//
//        jTable1.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable1));
//        jTable1.getColumn("comentario").setCellEditor(new CustomEditor(jTable1));
//
//        jTable1.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable1));
//        jTable1.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable1));
//
//        jTable1.getColumn("estado").setCellRenderer(new CustomRenderer1("REVISAR?"));
//        jTable1.getColumn("estado").setCellEditor(new CustomEditor1("REVISAR?", jTable1));
        suma();
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        // TODO add your handling code here:
        try {
            int selectedIndex = jTabbedPane1.getSelectedIndex();
            switch (selectedIndex) {
                case 0:
                    exportDataToExcel(jTable1);
                    break;
                case 1:
                    exportDataToExcel(jTable2);
                    break;
                case 2:
                    exportDataToExcel(jTable4);
                    break;
                case 3:
                    exportDataToExcel(jTable5);
                    break;
                case 4:
                    exportDataToExcel(jTable4);
                    break;
                default:
                    break;
            }
        } catch (IOException ex) {
        }
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void jMenuItem2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem2ActionPerformed
        // TODO add your handling code here:
        Scanner entrada = null;
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.showOpenDialog(fileChooser);
        try {
            String ruta = fileChooser.getSelectedFile().getAbsolutePath();
            System.out.println(ruta);
            File f = new File(ruta);

            CompletableFuture.runAsync(() -> {
                try {
                    jTabbedPane1.setVisible(false);
                    jPanel2.setVisible(false);
                    jPanel3.setVisible(false);
                    jPanel4.setVisible(false);

//                    cg = new VentanaCargar();
//                    cg.setAlwaysOnTop(true);
//                    cg.setUndecorated(true);
//                    cg.setLocationRelativeTo(this);
//                    cg.setVisible(true);
                    leerExcel(f);
                } catch (IOException ex) {
                    Exceptions.printStackTrace(ex);

                    jTabbedPane1.setVisible(false);
                    jPanel2.setVisible(false);
                    jPanel3.setVisible(false);
                    jPanel4.setVisible(false);
                }
            }).thenRun(() -> {
                CompletableFuture.runAsync(() -> {
                    try {
////////////////////////////////////////////////////////////////////////////////     
                        manejo();
////////////////////////////////////////////////////////////////////////////////
                        try {
                            int size = arrCruzeOK.size();
                            System.out.println(size);
                            int size1 = arrCruzeBAD.size();
                            System.out.println(size1);
                            int size2 = arrCruzeMuyMALA.size();
                            System.out.println(size2);

                            cargarTablas();
                            canFindAllFridaysInRange(jTable1, 1);
                        } catch (ParseException ex) {
                            Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                        }

                        jLabel2.setText("" + jTable1.getRowCount());

                        llenaTabla4();

                        suma();
                        selectedTabla4();
                        corrigueSelected();
////////////////////////////////////////////////////////////////////////////////
                    } catch (ParseException ex) {
                        Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (SQLException ex) {
                        Logger.getLogger(NewJFrame.class.getName()).log(Level.SEVERE, null, ex);
                    } finally {
                        jTabbedPane1.setVisible(true);
                        jPanel2.setVisible(true);
                        jPanel3.setVisible(true);
                        jPanel4.setVisible(true);
                    }
                }).thenRun(() -> {
                    jTabbedPane1.setVisible(true);
                    jPanel2.setVisible(true);
                    jPanel3.setVisible(true);
                    jPanel4.setVisible(true);
                });
            });
            entrada = new Scanner(f);
            while (entrada.hasNext()) {
                System.out.println(entrada.nextLine());
            }
        } catch (FileNotFoundException e) {
            Logger.getLogger(CustomEditor.class.getName()).log(Level.SEVERE, null, e);
        } catch (NullPointerException e) {
            Logger.getLogger(CustomEditor.class.getName()).log(Level.SEVERE, null, e);
        } catch (Exception e) {
            Logger.getLogger(CustomEditor.class.getName()).log(Level.SEVERE, null, e);
        } finally {
            if (entrada != null) {
                entrada.close();
                jTabbedPane1.setVisible(true);
                jPanel2.setVisible(true);
                jPanel3.setVisible(true);
                jPanel4.setVisible(true);
            }
        }
    }//GEN-LAST:event_jMenuItem2ActionPerformed

    private void jButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton8ActionPerformed
        // TODO add your handling code here:
        CompletableFuture.runAsync(() -> {
            opcion = false;
            File file = new File(System.getProperty("user.dir") + "\\config.ini");

            File directoryPath = new File(System.getProperty("user.dir"));
            String contents[] = directoryPath.list();
            String nombre = "";
            for (int i = 0; i < contents.length; i++) {
                System.out.println(contents[i]);
                if (contents[i].contains(".jar")) {
                    nombre = contents[i];
                    break;
                }
            }

            System.out.println("nombre " + nombre);

            try {
                byte[] bytes = null;
                int length = 0;
                int cont = 0;
                while (length == 0 && cont < 20) {
                    try {
                        bytes = consultaVersionDAO.ConsultaPrograma(v, nombre);
                        length = bytes.length;
                    } catch (Exception ex) {
                        length = 0;
                    }
                    System.out.println("length " + length);
                    System.out.println("cont " + cont);
                    Thread.sleep(1000);
                    cont++;
                }

                if (length != 0) {
                    System.out.println("1");
                    cont = 0;
                    try {
                        Ini ini = new Ini(file);
                        ini.get("Propiedades").put("box", 1);
                        ini.get("Version").put("v", v);
                        ini.store();
                        new ProcessBuilder("cmd", "/c start /min " + System.getProperty("user.dir") + "\\" + nombre + " ^& exit").start();
                        System.exit(0);
                    } catch (FileNotFoundException ex) {
                        JOptionPane.showMessageDialog(null, "Ups", ex.toString(), JOptionPane.ERROR_MESSAGE);
                        Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, "Ups", ex.toString(), JOptionPane.ERROR_MESSAGE);
                        Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                    }
                } else {
                    System.out.println("2");
                    cont = 0;
                    try {
                        Ini ini = new Ini(file);
                        ini.get("Propiedades").put("box", 1);
                        ini.get("Version").put("v", (Integer.valueOf(v) - 1));
                        ini.store();
                    } catch (FileNotFoundException ex) {
                        JOptionPane.showMessageDialog(null, "Ups", ex.toString(), JOptionPane.ERROR_MESSAGE);
                        Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, "Ups", ex.toString(), JOptionPane.ERROR_MESSAGE);
                        Logger.getLogger(Calendar.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            } catch (InterruptedException ex) {
                JOptionPane.showMessageDialog(null, "Ups", ex.toString(), JOptionPane.ERROR_MESSAGE);
                Logger.getLogger(java.util.Calendar.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
    }//GEN-LAST:event_jButton8ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        // TODO add your handling code here:
        int rowCount = jTable1.getRowCount();
        int rowCount1 = jTable2.getRowCount();
        int rowCount2 = jTable3.getRowCount();
        int rowCount3 = jTable5.getRowCount();
        int rowCount4 = jTable6.getRowCount();

        String proveedor = jComboBox3.getSelectedItem().toString();

        if (proveedor.equals("TODOS")) {
            ArrayList<Estadistica> arrEstadistica = new ArrayList<>();

            for (int i = 0; i < rowCount; i++) {
                String toString = jTable1.getValueAt(i, 7).toString();

//                if (proveedor.equals(toString)) {
                Estadistica estadistica = new Estadistica();

                String fechaEmision = jTable1.getValueAt(i, 8).toString();
                String saldo = jTable1.getValueAt(i, 10).toString();
                String notaCredito = jTable1.getValueAt(i, 11).toString();

                estadistica.setNombre(proveedor);
                estadistica.setFechaEmision(fechaEmision);
                estadistica.setSaldo(saldo);
                estadistica.setNotaCredito(notaCredito);

                arrEstadistica.add(estadistica);
//              }
            }

            for (int i = 0; i < rowCount1; i++) {
                String toString = jTable2.getValueAt(i, 7).toString();
//                if (proveedor.equals(toString)) {
                Estadistica estadistica = new Estadistica();

                String fechaEmision = jTable2.getValueAt(i, 8).toString();
                String saldo = jTable2.getValueAt(i, 10).toString();
                String notaCredito = jTable2.getValueAt(i, 11).toString();

                estadistica.setNombre(proveedor);
                estadistica.setFechaEmision(fechaEmision);
                estadistica.setSaldo(saldo);
                estadistica.setNotaCredito(notaCredito);

                arrEstadistica.add(estadistica);
//                }
            }

            for (int i = 0; i < rowCount2; i++) {
                String toString = jTable3.getValueAt(i, 7).toString();
//                if (proveedor.equals(toString)) {
                Estadistica estadistica = new Estadistica();

                String fechaEmision = jTable3.getValueAt(i, 8).toString();
                String saldo = jTable3.getValueAt(i, 10).toString();
                String notaCredito = jTable3.getValueAt(i, 11).toString();

                estadistica.setNombre(proveedor);
                estadistica.setFechaEmision(fechaEmision);
                estadistica.setSaldo(saldo);
                estadistica.setNotaCredito(notaCredito);

                arrEstadistica.add(estadistica);
//                }
            }

            for (int i = 0; i < rowCount3; i++) {
                String toString = jTable5.getValueAt(i, 7).toString();
//                if (proveedor.equals(toString)) {
                Estadistica estadistica = new Estadistica();

                String fechaEmision = jTable5.getValueAt(i, 8).toString();
                String saldo = jTable5.getValueAt(i, 10).toString();
                String notaCredito = jTable5.getValueAt(i, 11).toString();

                estadistica.setNombre(proveedor);
                estadistica.setFechaEmision(fechaEmision);
                estadistica.setSaldo(saldo);
                estadistica.setNotaCredito(notaCredito);

                arrEstadistica.add(estadistica);
//                }
            }

            for (int i = 0; i < rowCount4; i++) {
                String toString = jTable6.getValueAt(i, 7).toString();
//                if (proveedor.equals(toString)) {
                Estadistica estadistica = new Estadistica();

                String fechaEmision = jTable6.getValueAt(i, 8).toString();
                String saldo = jTable6.getValueAt(i, 10).toString();
                String notaCredito = jTable6.getValueAt(i, 11).toString();

                estadistica.setNombre(proveedor);
                estadistica.setFechaEmision(fechaEmision);
                estadistica.setSaldo(saldo);
                estadistica.setNotaCredito(notaCredito);

                arrEstadistica.add(estadistica);
//                }
            }

            ArrayList<String> arrFechasEmision = new ArrayList<>();

            arrEstadistica.stream().forEach((Estadistica estadistica) -> {
                String fechaEmision = estadistica.getFechaEmision();
                arrFechasEmision.add(fechaEmision);
            });

            Set<String> hashSet = new HashSet<String>(arrFechasEmision);
            arrFechasEmision.clear();
            arrFechasEmision.addAll(hashSet);

            System.out.println(arrFechasEmision);

            ArrayList<String> arrAños = new ArrayList<>();
            arrFechasEmision.stream().forEach((var fechas) -> {
                String[] split = fechas.split("/");
                arrAños.add(split[2]);
            });

            hashSet = new HashSet<String>(arrAños);
            arrAños.clear();
            arrAños.addAll(hashSet);

            System.out.println(arrAños);

            ArrayList<ArrayList<String>> arrMesesPorAño = new ArrayList<>();

            arrAños.stream().forEach((var año) -> {
                ArrayList<String> arrFechas = new ArrayList<>();
                arrFechasEmision.stream().forEach((var fecha) -> {
                    if (fecha.contains("/" + año)) {
                        arrFechas.add(fecha.substring(2));
                    }
                });
                ArrayList<String> regresaSinDuplicados = regresaSinDuplicados(arrFechas);
                arrMesesPorAño.add(regresaSinDuplicados);
            });

            ArrayList<AñoMesSaldo> arrAñoMesSaldo = new ArrayList<>();
            arrMesesPorAño.stream().forEach((ArrayList<String> arrFecha) -> {
                System.out.println(arrFecha);
                System.out.println("///////");

                arrFecha.stream().forEach((String meses) -> {
                    AñoMesSaldo añoMesSaldo = new AñoMesSaldo();
                    long sumaSaldos = 0;
                    long sumaNotaCredito = 0;
                    String añox = "";
                    String mes = "";

                    for (int i = 0; i < arrEstadistica.size(); i++) {
                        Estadistica estadistica = arrEstadistica.get(i);
                        String fechaEmision = estadistica.getFechaEmision();

                        if (fechaEmision.contains(meses)) {
                            String saldo = estadistica.getSaldo();
                            String notaCredito = estadistica.getNotaCredito();

//                            System.out.println("saldo " + saldo);
//                            System.out.println("notaCredito " + notaCredito);
                            sumaSaldos = sumaSaldos + Long.valueOf(saldo.replace(".", ""));
                            sumaNotaCredito = sumaNotaCredito + Long.valueOf(notaCredito);

                            if (meses.equals("/08/2021")) {
                                System.out.println("saldo " + saldo);
                                System.out.println("sumaSaldos " + sumaSaldos);
                            }

                            for (int j = 0; j < arrAños.size(); j++) {
                                String año = arrAños.get(j);

                                if (añox.equals("")) {
                                    if (estadistica.getFechaEmision().contains(año)) {
                                        añox = año;
                                        break;
                                    }
                                }
                            }

                            if (mes.equals("")) {
                                String[] split = meses.split("/");
                                mes = split[1];
                            }
                        }
                    }
                    añoMesSaldo.setAño(añox);
                    añoMesSaldo.setMes(mes);
                    añoMesSaldo.setSaldo("" + sumaSaldos);
                    añoMesSaldo.setNotaCredito("" + sumaNotaCredito);
                    arrAñoMesSaldo.add(añoMesSaldo);
                    System.out.println("---------");
                });
            });

            DefaultCategoryDataset dataset = new DefaultCategoryDataset();

            ArrayList<ArrayList<Long>> saldosPorAño = new ArrayList<>();

            for (int j = 0; j < arrAños.size(); j++) {
                String get = arrAños.get(j);
                ArrayList<Long> arr = new ArrayList<>();

                String enero = "";
                String febrero = "";
                String marzo = "";
                String abril = "";
                String mayo = "";
                String junio = "";
                String julio = "";
                String agosto = "";
                String septiembre = "";
                String octubre = "";
                String noviembre = "";
                String diciembre = "";

                for (int i = 0; i < arrAñoMesSaldo.size(); i++) {
                    AñoMesSaldo añoMesSaldo = arrAñoMesSaldo.get(i);

                    String año = añoMesSaldo.getAño();
                    String mes = añoMesSaldo.getMes();
                    String saldo = añoMesSaldo.getSaldo();
                    String notaCredito = añoMesSaldo.getNotaCredito();

                    if (año.equals(get)) {
                        System.out.println("mes " + mes);
                        System.out.println("saldo " + saldo);
                        System.out.println("notaCredito " + notaCredito);

                        if (mes.equals("01")) {
                            enero = saldo;
                        } else if (mes.equals("02")) {
                            febrero = saldo;
                        } else if (mes.equals("03")) {
                            marzo = saldo;
                        } else if (mes.equals("04")) {
                            abril = saldo;
                        } else if (mes.equals("05")) {
                            mayo = saldo;
                        } else if (mes.equals("06")) {
                            junio = saldo;
                        } else if (mes.equals("07")) {
                            julio = saldo;
                        } else if (mes.equals("08")) {
                            agosto = saldo;
                        } else if (mes.equals("09")) {
                            septiembre = saldo;
                        } else if (mes.equals("10")) {
                            octubre = saldo;
                        } else if (mes.equals("11")) {
                            noviembre = saldo;
                        } else if (mes.equals("12")) {
                            diciembre = saldo;
                        }
                    }
                }

                try {
                    arr.add(Long.valueOf(enero));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(febrero));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(marzo));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(abril));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(mayo));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(junio));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(julio));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(agosto));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(septiembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(octubre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));;
                }
                try {
                    arr.add(Long.valueOf(noviembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(diciembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }

                saldosPorAño.add(arr);
            }

            for (int i = 0; i < saldosPorAño.size(); i++) {
                ArrayList<Long> get = saldosPorAño.get(i);
                String año = arrAños.get(i);

                for (int j = 0; j < get.size(); j++) {
                    long get1 = get.get(j);

                    if (j == 0) {
                        dataset.setValue(get1, año, "Enero");
                    }
                    if (j == 1) {
                        dataset.setValue(get1, año, "Febrero");
                    }
                    if (j == 2) {
                        dataset.setValue(get1, año, "Marzo");
                    }
                    if (j == 3) {
                        dataset.setValue(get1, año, "Abril");
                    }
                    if (j == 4) {
                        dataset.setValue(get1, año, "Mayo");
                    }
                    if (j == 5) {
                        dataset.setValue(get1, año, "Junio");
                    }
                    if (j == 6) {
                        dataset.setValue(get1, año, "Julio");
                    }
                    if (j == 7) {
                        dataset.setValue(get1, año, "Agosto");
                    }
                    if (j == 8) {
                        dataset.setValue(get1, año, "Septiembre");
                    }
                    if (j == 9) {
                        dataset.setValue(get1, año, "Octubre");
                    }
                    if (j == 10) {
                        dataset.setValue(get1, año, "Noviembre");
                    }
                    if (j == 11) {
                        dataset.setValue(get1, año, "Diciembre");
                    }
                }
            }

            JFreeChart chart = ChartFactory.createBarChart("Saldos Por Mes", "Meses", "Saldos", dataset, PlotOrientation.VERTICAL, rootPaneCheckingEnabled, rootPaneCheckingEnabled, rootPaneCheckingEnabled);
            CategoryPlot p = chart.getCategoryPlot();
            ((BarRenderer) p.getRenderer()).setBarPainter(new StandardBarPainter());
            p.setRangeGridlinePaint(Color.black);

            try {
                jPanel6.remove(0);
            } catch (Exception ex) {

            }

            jPanel6.setLayout(new java.awt.BorderLayout());
            ChartPanel CP = new ChartPanel(chart);
            jPanel6.add(CP, BorderLayout.CENTER);
            jPanel6.validate();
        } else {
            ArrayList<Estadistica> arrEstadistica = new ArrayList<>();

            for (int i = 0; i < rowCount; i++) {
                String toString = jTable1.getValueAt(i, 7).toString();

                if (proveedor.equals(toString)) {
                    Estadistica estadistica = new Estadistica();

                    String fechaEmision = jTable1.getValueAt(i, 8).toString();
                    String saldo = jTable1.getValueAt(i, 10).toString();
                    String notaCredito = jTable1.getValueAt(i, 11).toString();

                    estadistica.setNombre(proveedor);
                    estadistica.setFechaEmision(fechaEmision);
                    estadistica.setSaldo(saldo);
                    estadistica.setNotaCredito(notaCredito);

                    arrEstadistica.add(estadistica);
                }
            }

            for (int i = 0; i < rowCount1; i++) {
                String toString = jTable2.getValueAt(i, 7).toString();
                if (proveedor.equals(toString)) {
                    Estadistica estadistica = new Estadistica();

                    String fechaEmision = jTable2.getValueAt(i, 8).toString();
                    String saldo = jTable2.getValueAt(i, 10).toString();
                    String notaCredito = jTable2.getValueAt(i, 11).toString();

                    estadistica.setNombre(proveedor);
                    estadistica.setFechaEmision(fechaEmision);
                    estadistica.setSaldo(saldo);
                    estadistica.setNotaCredito(notaCredito);

                    arrEstadistica.add(estadistica);
                }
            }

            for (int i = 0; i < rowCount2; i++) {
                String toString = jTable3.getValueAt(i, 7).toString();
                if (proveedor.equals(toString)) {
                    Estadistica estadistica = new Estadistica();

                    String fechaEmision = jTable3.getValueAt(i, 8).toString();
                    String saldo = jTable3.getValueAt(i, 10).toString();
                    String notaCredito = jTable3.getValueAt(i, 11).toString();

                    estadistica.setNombre(proveedor);
                    estadistica.setFechaEmision(fechaEmision);
                    estadistica.setSaldo(saldo);
                    estadistica.setNotaCredito(notaCredito);

                    arrEstadistica.add(estadistica);
                }
            }

            for (int i = 0; i < rowCount3; i++) {
                String toString = jTable5.getValueAt(i, 7).toString();
                if (proveedor.equals(toString)) {
                    Estadistica estadistica = new Estadistica();

                    String fechaEmision = jTable5.getValueAt(i, 8).toString();
                    String saldo = jTable5.getValueAt(i, 10).toString();
                    String notaCredito = jTable5.getValueAt(i, 11).toString();

                    estadistica.setNombre(proveedor);
                    estadistica.setFechaEmision(fechaEmision);
                    estadistica.setSaldo(saldo);
                    estadistica.setNotaCredito(notaCredito);

                    arrEstadistica.add(estadistica);
                }
            }

            for (int i = 0; i < rowCount4; i++) {
                String toString = jTable6.getValueAt(i, 7).toString();
                if (proveedor.equals(toString)) {
                    Estadistica estadistica = new Estadistica();

                    String fechaEmision = jTable6.getValueAt(i, 8).toString();
                    String saldo = jTable6.getValueAt(i, 10).toString();
                    String notaCredito = jTable6.getValueAt(i, 11).toString();

                    estadistica.setNombre(proveedor);
                    estadistica.setFechaEmision(fechaEmision);
                    estadistica.setSaldo(saldo);
                    estadistica.setNotaCredito(notaCredito);

                    arrEstadistica.add(estadistica);
                }
            }

            ArrayList<String> arrFechasEmision = new ArrayList<>();

            arrEstadistica.stream().forEach((Estadistica estadistica) -> {
                String fechaEmision = estadistica.getFechaEmision();
                arrFechasEmision.add(fechaEmision);
            });

            Set<String> hashSet = new HashSet<String>(arrFechasEmision);
            arrFechasEmision.clear();
            arrFechasEmision.addAll(hashSet);

            System.out.println(arrFechasEmision);

            ArrayList<String> arrAños = new ArrayList<>();
            arrFechasEmision.stream().forEach((var fechas) -> {
                String[] split = fechas.split("/");
                arrAños.add(split[2]);
            });

            hashSet = new HashSet<String>(arrAños);
            arrAños.clear();
            arrAños.addAll(hashSet);

            System.out.println(arrAños);

            ArrayList<ArrayList<String>> arrMesesPorAño = new ArrayList<>();

            arrAños.stream().forEach((var año) -> {
                ArrayList<String> arrFechas = new ArrayList<>();
                arrFechasEmision.stream().forEach((var fecha) -> {
                    if (fecha.contains("/" + año)) {
                        arrFechas.add(fecha.substring(2));
                    }
                });
                ArrayList<String> regresaSinDuplicados = regresaSinDuplicados(arrFechas);
                arrMesesPorAño.add(regresaSinDuplicados);
            });

            ArrayList<AñoMesSaldo> arrAñoMesSaldo = new ArrayList<>();
            arrMesesPorAño.stream().forEach((ArrayList<String> arrFecha) -> {
                System.out.println(arrFecha);
                System.out.println("///////");

                arrFecha.stream().forEach((String meses) -> {
                    AñoMesSaldo añoMesSaldo = new AñoMesSaldo();
                    long sumaSaldos = 0;
                    long sumaNotaCredito = 0;
                    String añox = "";
                    String mes = "";

                    for (int i = 0; i < arrEstadistica.size(); i++) {
                        Estadistica estadistica = arrEstadistica.get(i);
                        String fechaEmision = estadistica.getFechaEmision();

                        if (fechaEmision.contains(meses)) {
                            String saldo = estadistica.getSaldo();
                            String notaCredito = estadistica.getNotaCredito();

                            System.out.println("saldo " + saldo);
                            System.out.println("notaCredito " + notaCredito);

                            sumaSaldos = sumaSaldos + Long.valueOf(saldo.replace(".", ""));
                            sumaNotaCredito = sumaNotaCredito + Long.valueOf(notaCredito);

                            for (int j = 0; j < arrAños.size(); j++) {
                                String año = arrAños.get(j);

                                if (añox.equals("")) {
                                    if (estadistica.getFechaEmision().contains(año)) {
                                        añox = año;
                                        break;
                                    }
                                }
                            }

                            if (mes.equals("")) {
                                String[] split = meses.split("/");
                                mes = split[1];
                            }
                        }
                    }
                    añoMesSaldo.setAño(añox);
                    añoMesSaldo.setMes(mes);
                    añoMesSaldo.setSaldo("" + sumaSaldos);
                    añoMesSaldo.setNotaCredito("" + sumaNotaCredito);
                    arrAñoMesSaldo.add(añoMesSaldo);
                    System.out.println("---------");
                });
            });

            DefaultCategoryDataset dataset = new DefaultCategoryDataset();

            ArrayList<ArrayList<Long>> saldosPorAño = new ArrayList<>();

            for (int j = 0; j < arrAños.size(); j++) {
                String get = arrAños.get(j);
                ArrayList<Long> arr = new ArrayList<>();

                String enero = "";
                String febrero = "";
                String marzo = "";
                String abril = "";
                String mayo = "";
                String junio = "";
                String julio = "";
                String agosto = "";
                String septiembre = "";
                String octubre = "";
                String noviembre = "";
                String diciembre = "";

                for (int i = 0; i < arrAñoMesSaldo.size(); i++) {
                    AñoMesSaldo añoMesSaldo = arrAñoMesSaldo.get(i);

                    String año = añoMesSaldo.getAño();
                    String mes = añoMesSaldo.getMes();
                    String saldo = añoMesSaldo.getSaldo();
                    String notaCredito = añoMesSaldo.getNotaCredito();

                    if (año.equals(get)) {
                        System.out.println("mes " + mes);
                        System.out.println("saldo " + saldo);
                        System.out.println("notaCredito " + notaCredito);

                        if (mes.equals("01")) {
                            enero = saldo;
                        } else if (mes.equals("02")) {
                            febrero = saldo;
                        } else if (mes.equals("03")) {
                            marzo = saldo;
                        } else if (mes.equals("04")) {
                            abril = saldo;
                        } else if (mes.equals("05")) {
                            mayo = saldo;
                        } else if (mes.equals("06")) {
                            junio = saldo;
                        } else if (mes.equals("07")) {
                            julio = saldo;
                        } else if (mes.equals("08")) {
                            agosto = saldo;
                        } else if (mes.equals("09")) {
                            septiembre = saldo;
                        } else if (mes.equals("10")) {
                            octubre = saldo;
                        } else if (mes.equals("11")) {
                            noviembre = saldo;
                        } else if (mes.equals("12")) {
                            diciembre = saldo;
                        }
                    }
                }

                try {
                    arr.add(Long.valueOf(enero));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(febrero));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(marzo));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(abril));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(mayo));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(junio));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(julio));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(agosto));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(septiembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(octubre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(noviembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }
                try {
                    arr.add(Long.valueOf(diciembre));
                } catch (Exception ex) {
                    arr.add(Long.valueOf(0));
                }

                saldosPorAño.add(arr);
            }

            for (int i = 0; i < saldosPorAño.size(); i++) {
                ArrayList<Long> get = saldosPorAño.get(i);
                String año = arrAños.get(i);

                for (int j = 0; j < get.size(); j++) {
                    long get1 = get.get(j);

                    if (j == 0) {
                        dataset.setValue(get1, año, "Enero");
                    }
                    if (j == 1) {
                        dataset.setValue(get1, año, "Febrero");
                    }
                    if (j == 2) {
                        dataset.setValue(get1, año, "Marzo");
                    }
                    if (j == 3) {
                        dataset.setValue(get1, año, "Abril");
                    }
                    if (j == 4) {
                        dataset.setValue(get1, año, "Mayo");
                    }
                    if (j == 5) {
                        dataset.setValue(get1, año, "Junio");
                    }
                    if (j == 6) {
                        dataset.setValue(get1, año, "Julio");
                    }
                    if (j == 7) {
                        dataset.setValue(get1, año, "Agosto");
                    }
                    if (j == 8) {
                        dataset.setValue(get1, año, "Septiembre");
                    }
                    if (j == 9) {
                        dataset.setValue(get1, año, "Octubre");
                    }
                    if (j == 10) {
                        dataset.setValue(get1, año, "Noviembre");
                    }
                    if (j == 11) {
                        dataset.setValue(get1, año, "Diciembre");
                    }
                }
            }

            JFreeChart chart = ChartFactory.createBarChart("Saldos Por Mes", "Meses", "Saldos", dataset, PlotOrientation.VERTICAL, rootPaneCheckingEnabled, rootPaneCheckingEnabled, rootPaneCheckingEnabled);
            CategoryPlot p = chart.getCategoryPlot();
            ((BarRenderer) p.getRenderer()).setBarPainter(new StandardBarPainter());
            p.setRangeGridlinePaint(Color.black);

            try {
                jPanel6.remove(0);
            } catch (Exception ex) {

            }

            jPanel6.setLayout(new java.awt.BorderLayout());
            ChartPanel CP = new ChartPanel(chart);
            jPanel6.add(CP, BorderLayout.CENTER);
            jPanel6.validate();
        }
    }//GEN-LAST:event_jButton3ActionPerformed

    public static ArrayList<String> regresaSinDuplicados(ArrayList<String> arr) {

        Set<String> hashSet = new HashSet<String>(arr);
        arr.clear();
        arr.addAll(hashSet);

        return arr;
    }

    public void exportDataToExcel(JTable t) throws IOException {
        JFileChooser chooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls");
        chooser.setFileFilter(filter);
        chooser.setDialogTitle("Guardar archivo");
        chooser.setAcceptAllFileFilterUsed(false);
        if (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
            String ruta = chooser.getSelectedFile().toString().concat(".xls");
            try {
                File archivoXLS = new File(ruta);
                if (archivoXLS.exists()) {
                    archivoXLS.delete();
                }
                archivoXLS.createNewFile();
                Workbook libro = new HSSFWorkbook();
                FileOutputStream archivo = new FileOutputStream(archivoXLS);
                Sheet hoja = libro.createSheet("Mi hoja de trabajo 1");
                hoja.setDisplayGridlines(false);

                for (int f = 0; f < t.getRowCount(); f++) {
                    Row fila = hoja.createRow(f);
                    for (int c = 0; c < t.getColumnCount(); c++) {
                        Cell celda = fila.createCell(c);
                        if (f == 0) {
                            celda.setCellValue(t.getColumnName(c));
                        }
                    }
                }

                int filaInicio = 1;
                for (int f = 0; f < t.getRowCount(); f++) {
                    Row fila = hoja.createRow(filaInicio);
                    filaInicio++;
                    for (int c = 0; c < t.getColumnCount(); c++) {
                        Cell celda = fila.createCell(c);
                        if (t.getValueAt(f, c) instanceof Double) {
                            celda.setCellValue(Double.parseDouble(t.getValueAt(f, c).toString()));
                        } else if (t.getValueAt(f, c) instanceof Float) {
                            celda.setCellValue(Float.parseFloat((String) t.getValueAt(f, c)));
                        } else {
                            try {
                                String toString = t.getValueAt(f, c).toString();
                                DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                Date parse = dateFormat.parse(toString);

                                SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
                                String strDate = formatter.format(parse);
                                celda.setCellValue(strDate);
                            } catch (Exception ex) {
                                celda.setCellValue(String.valueOf(t.getValueAt(f, c)));
                            }
                        }
                    }
                }
                libro.write(archivo);
                archivo.close();
                Desktop.getDesktop().open(archivoXLS);
            } catch (IOException | NumberFormatException e) {
                throw e;
            }
        }
    } //aqui JaSON!

    public static void filter(Date startDate, Date endDate, JTable jTable) {
        System.out.println("1");
        Instant instant = startDate.toInstant();
        ZonedDateTime zdt = instant.atZone(ZoneId.systemDefault());
        LocalDate date = zdt.toLocalDate();
        LocalDate minusDays = date.minusDays(1);
        ZoneId defaultZoneId = ZoneId.systemDefault();
        Date date1 = Date.from(minusDays.atStartOfDay(defaultZoneId).toInstant());

        Instant instant1 = endDate.toInstant();
        ZonedDateTime zdt1 = instant1.atZone(ZoneId.systemDefault());
        LocalDate date2 = zdt1.toLocalDate();
        LocalDate plusDays = date2.plusDays(1);
        ZoneId defaultZoneId1 = ZoneId.systemDefault();
        Date date3 = Date.from(plusDays.atStartOfDay(defaultZoneId1).toInstant());

        List<RowFilter<Object, Object>> filters = new ArrayList<>(2);
        filters.add(RowFilter.dateFilter(ComparisonType.AFTER, date1, 9));
        filters.add(RowFilter.dateFilter(ComparisonType.BEFORE, date3, 9));

        if (!jComboBox2.getSelectedItem().toString().equals("TODOS")) {
            filters.add(RowFilter.regexFilter(jComboBox2.getSelectedItem().toString(), 6));
        }

        DefaultTableModel dtm = (DefaultTableModel) jTable.getModel();
        TableRowSorter<DefaultTableModel> tr = new TableRowSorter<>(dtm);
        jTable.setRowSorter(tr);
        RowFilter<Object, Object> rf = RowFilter.andFilter(filters);
        tr.setRowFilter(rf);
    }

    public static void cargarTablas() throws ParseException {

        DefaultTableModel model = new DefaultTableModel();

        model.addColumn("pkNumeroCuota");
        model.addColumn("numeroOrdenDoc");
        model.addColumn("guiaChilemat");
        model.addColumn("guiaProveedor");
        model.addColumn("local");
//        model.addColumn("numeroOrdenIngreso");
        model.addColumn("numeroGuia");
//        model.addColumn("fechaRecepcion");
        model.addColumn("sucursal");
        model.addColumn("proveedor");
        model.addColumn("fechaEmision");
        model.addColumn("fechaVencimiento");
//        model.addColumn("total");
        model.addColumn("saldo");
        model.addColumn("Nota de Credito");
        model.addColumn("totalNCuota");
        model.addColumn("procentaje");
        model.addColumn("comentario");
        model.addColumn("estado");

        Object[] fila = new Object[16];

//                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
//                sdf.setLenient(false);
        /*    
                String numeroOrdenDoc;
                String guiaChilemat;
                String guiaProveedor;
                String local;

                String numeroOrdenIngreso;
                String numeroGuia;
                String fechaRecepcion;
                String sucursal;

                String proveedor;
                String fechaEmision;
                String total;
                String saldo;

                String totalNCuota;
                String procentaje;
         */
        for (int i = 0; i < arrCruzeOK.size(); ++i) {
            System.out.println("--------------------------------------------> " + i);
            Cruze get = arrCruzeOK.get(i);

            fila[0] = get.getPkNumeroCuota();
            fila[1] = get.getNumeroOrdenDoc();
            fila[2] = get.getGuiaChilemat();
            fila[3] = get.getGuiaProveedor();
            fila[4] = get.getLocal();

//           fila[5] = get.getNumeroOrdenIngreso();
            fila[5] = get.getNumeroGuia();
//            fila[6] = get.getFechaRecepcion();
            fila[6] = get.getSucursal();

            fila[7] = get.getProveedor();
            fila[8] = get.getFechaEmision();

            try {
                System.out.println("get.getFechaVencimiento() " + get.getFechaVencimiento());
                Date date = new SimpleDateFormat("dd/MM/yyyy").parse(get.getFechaVencimiento());
                fila[9] = date;
            } catch (Exception ex) {
                fila[9] = "";
            }

            try {
                System.out.println(get.getSaldo());
                double firstNumber = Double.valueOf(get.getSaldo());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[10] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[11] = String.valueOf(get.getComentarioNotaDeCredito());

            try {
                System.out.println(get.getTotalNCuota());
                double firstNumber = Double.valueOf(get.getTotalNCuota());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[12] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[13] = get.getProcentaje();

            String comnetario = get.getComnetario();
            System.out.println("---------------------------------------------> comnetario " + comnetario);
            if (comnetario != null) {
                fila[14] = comnetario;
            } else {
                fila[14] = "";
            }

            model.addRow(fila);
        }

        jTable1.setModel(model);

        TableRowSorter trs = new TableRowSorter(model);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

//        trs.setComparator(12, new IntComparator());
//        trs.setComparator(2, new StringComparator2());
//        trs.setComparator(3, new IntComparator());
//        trs.setComparator(4, new StringComparator());
//        trs.setComparator(5, new StringComparator());
//        trs.setComparator(6, new StringComparator());
        jTable1.setRowSorter(trs);
//                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
//                renderer.setHorizontalAlignment(JLabel.CENTER);

        DefaultTableCellRenderer centerRenderer2 = new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                try {
                    DateFormat dateFormat0 = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                    Date parse = dateFormat0.parse(value.toString());
                    DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                    String strDate = dateFormat.format(parse);
                    value = strDate;
                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                } catch (ParseException ex) {
                    System.out.println(ex);
                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                }
            }
        };
        centerRenderer2.setHorizontalAlignment(JLabel.CENTER);

        DefaultTableCellRenderer Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

//        jTable1.getColumnModel().getColumn(0).setWidth(0);
//        jTable1.getColumnModel().getColumn(0).setMinWidth(0);
//        jTable1.getColumnModel().getColumn(0).setMaxWidth(0);
        jTable1.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(5).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(6).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(7).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(8).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
        jTable1.getColumnModel().getColumn(10).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(11).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(12).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(13).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(14).setCellRenderer(Renderer);
        jTable1.getColumnModel().getColumn(15).setCellRenderer(Renderer);

        jTable1.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);

        jTable1.getColumnModel().getColumn(1).setPreferredWidth(30);
        jTable1.getColumnModel().getColumn(2).setPreferredWidth(30);
        jTable1.getColumnModel().getColumn(3).setPreferredWidth(30);
        jTable1.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable1.getColumnModel().getColumn(11).setPreferredWidth(150);
        jTable1.getColumnModel().getColumn(14).setPreferredWidth(250);
        jTable1.getColumnModel().getColumn(15).setPreferredWidth(75);

        jTable1.setRowHeight(80);
        jTable1.setShowHorizontalLines(true);
        jTable1.setShowVerticalLines(true);

        jTable1.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable1));
        jTable1.getColumn("comentario").setCellEditor(new CustomEditor(jTable1));

        jTable1.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable1));
        jTable1.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable1));

        jTable1.getColumn("estado").setCellRenderer(new CustomRenderer1("REVISAR?"));
        jTable1.getColumn("estado").setCellEditor(new CustomEditor1("REVISAR?", jTable1));
////////////////////////////////////////////////////////////////////////////////
        model = new DefaultTableModel();
        model.addColumn("pkNumeroCuota");
        model.addColumn("numeroOrdenDoc");
        model.addColumn("guiaChilemat");
        model.addColumn("guiaProveedor");
        model.addColumn("local");
//        model.addColumn("numeroOrdenIngreso");
        model.addColumn("numeroGuia");
//        model.addColumn("fechaRecepcion");
        model.addColumn("sucursal");
        model.addColumn("proveedor");
        model.addColumn("fechaEmision");
        model.addColumn("fechaVencimiento");
//        model.addColumn("total");
        model.addColumn("saldo");
        model.addColumn("Nota de Credito");
        model.addColumn("totalNCuota");
        model.addColumn("procentaje");
        model.addColumn("comentario");
        model.addColumn("estado");
        model.addColumn("estado2");
        model.addColumn("estado3");

        fila = new Object[18];

//                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
//                sdf.setLenient(false);
        /*    
                String numeroOrdenDoc;
                String guiaChilemat;
                String guiaProveedor;
                String local;

                String numeroOrdenIngreso;
                String numeroGuia;
                String fechaRecepcion;
                String sucursal;

                String proveedor;
                String fechaEmision;
                String total;
                String saldo;

                String totalNCuota;
                String procentaje;
         */
        for (int i = 0; i < arrCruzeBAD.size(); ++i) {
            System.out.println("--------------------------------------------> " + i);
            Cruze get = arrCruzeBAD.get(i);

            fila[0] = get.getPkNumeroCuota();
            fila[1] = get.getNumeroOrdenDoc();
            fila[2] = get.getGuiaChilemat();
            fila[3] = get.getGuiaProveedor();
            fila[4] = get.getLocal();

//            fila[5] = get.getNumeroOrdenIngreso();
            fila[5] = get.getNumeroGuia();
//            fila[6] = get.getFechaRecepcion();
            fila[6] = get.getSucursal();

            fila[7] = get.getProveedor();
            fila[8] = get.getFechaEmision();

            try {
                System.out.println("get.getFechaVencimiento() " + get.getFechaVencimiento());
                Date date = new SimpleDateFormat("dd/MM/yyyy").parse(get.getFechaVencimiento());
                fila[9] = date;
            } catch (Exception ex) {
                fila[9] = "";
            }

//            fila[10] = get.getTotal();
            try {
                System.out.println(get.getSaldo());
                double firstNumber = Double.valueOf(get.getSaldo());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[10] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[11] = String.valueOf(get.getComentarioNotaDeCredito());

            try {
                System.out.println(get.getTotalNCuota());
                double firstNumber = Double.valueOf(get.getTotalNCuota());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[12] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[13] = get.getProcentaje();

            String comnetario = get.getComnetario();
            System.out.println("---------------------------------------------> comnetario " + comnetario);
            if (comnetario != null) {
                fila[14] = comnetario;
            } else {
                fila[14] = "";
            }
            model.addRow(fila);
        }
        jTable2.setModel(model);

        trs = new TableRowSorter(model);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

//        trs.setComparator(3, new IntComparator());
//        trs.setComparator(4, new StringComparator());
//        trs.setComparator(5, new StringComparator());
//        trs.setComparator(6, new StringComparator());
        jTable2.setRowSorter(trs);
//                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
//                renderer.setHorizontalAlignment(JLabel.CENTER);
        Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

//        jTable2.getColumnModel().getColumn(0).setWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMinWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMaxWidth(0);
        jTable2.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(5).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(6).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(7).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(8).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
        jTable2.getColumnModel().getColumn(10).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(11).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(12).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(13).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(14).setCellRenderer(Renderer);
        jTable2.getColumnModel().getColumn(15).setCellRenderer(Renderer);

        jTable2.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);

        jTable2.getColumnModel().getColumn(1).setPreferredWidth(30);
        jTable2.getColumnModel().getColumn(2).setPreferredWidth(30);
        jTable2.getColumnModel().getColumn(3).setPreferredWidth(30);
        jTable2.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable2.getColumnModel().getColumn(11).setPreferredWidth(150);
        jTable2.getColumnModel().getColumn(14).setPreferredWidth(250);
        jTable2.getColumnModel().getColumn(15).setPreferredWidth(75);
        jTable2.getColumnModel().getColumn(16).setPreferredWidth(75);

        /*
            jTable3.getColumnModel().getColumn(1).setPreferredWidth(30);
            jTable3.getColumnModel().getColumn(2).setPreferredWidth(30);
            jTable3.getColumnModel().getColumn(3).setPreferredWidth(30);
            jTable3.getColumnModel().getColumn(4).setPreferredWidth(5);
            jTable3.getColumnModel().getColumn(11).setPreferredWidth(150);
            jTable3.getColumnModel().getColumn(14).setPreferredWidth(250);
            jTable3.getColumnModel().getColumn(15).setPreferredWidth(75);
         */
        jTable2.setRowHeight(80);
        jTable2.setShowHorizontalLines(true);
        jTable2.setShowVerticalLines(true);

        jTable2.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable2));
        jTable2.getColumn("comentario").setCellEditor(new CustomEditor(jTable2));

        jTable2.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable2));
        jTable2.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable2));

        jTable2.getColumn("estado").setCellRenderer(new CustomRenderer1("OK?"));
        jTable2.getColumn("estado").setCellEditor(new CustomEditor1("OK?", jTable2));

        jTable2.getColumn("estado2").setCellRenderer(new CustomRenderer1("MALA?"));
        jTable2.getColumn("estado2").setCellEditor(new CustomEditor1("MALA?", jTable2));

        jTable2.getColumn("estado3").setCellRenderer(new CustomRenderer1("OC?"));
        jTable2.getColumn("estado3").setCellEditor(new CustomEditor1("OC?", jTable2));
        ////////////////////////////////////////////////////////////////////////

        model = new DefaultTableModel();
        model.addColumn("pkNumeroCuota");
        model.addColumn("numeroOrdenDoc");
        model.addColumn("guiaChilemat");
        model.addColumn("guiaProveedor");
        model.addColumn("local");
//        model.addColumn("numeroOrdenIngreso");
        model.addColumn("numeroGuia");
//        model.addColumn("fechaRecepcion");
        model.addColumn("sucursal");
        model.addColumn("proveedor");
        model.addColumn("fechaEmision");
        model.addColumn("fechaVencimiento");
//        model.addColumn("total");
        model.addColumn("saldo");
        model.addColumn("Nota de Credito");
        model.addColumn("totalNCuota");
        model.addColumn("procentaje");
        model.addColumn("comentario");
        model.addColumn("estado");
        model.addColumn("estado1");

        fila = new Object[16];

//                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
//                sdf.setLenient(false);
        /*    
                String numeroOrdenDoc;
                String guiaChilemat;
                String guiaProveedor;
                String local;

                String numeroOrdenIngreso;
                String numeroGuia;
                String fechaRecepcion;
                String sucursal;

                String proveedor;
                String fechaEmision;
                String total;
                String saldo;

                String totalNCuota;
                String procentaje;
         */
        for (int i = 0; i < arrCruzeMuyMALA.size(); ++i) {
            System.out.println("--------------------------------------------> " + i);
            Cruze get = arrCruzeMuyMALA.get(i);

            fila[0] = get.getPkNumeroCuota();
            fila[1] = get.getNumeroOrdenDoc();
            fila[2] = get.getGuiaChilemat();
            fila[3] = get.getGuiaProveedor();
            fila[4] = get.getLocal();

//            fila[5] = get.getNumeroOrdenIngreso();
            fila[5] = get.getNumeroGuia();
//            fila[6] = get.getFechaRecepcion();
            fila[6] = get.getSucursal();

            fila[7] = get.getProveedor();
            fila[8] = get.getFechaEmision();

            try {
                System.out.println("get.getFechaVencimiento() " + get.getFechaVencimiento());
                Date date = new SimpleDateFormat("dd/MM/yyyy").parse(get.getFechaVencimiento());
                fila[9] = date;
            } catch (Exception ex) {
                fila[9] = "";
            }

//            fila[10] = get.getTotal();
            try {
                System.out.println(get.getSaldo());
                double firstNumber = Double.valueOf(get.getSaldo());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[10] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[11] = String.valueOf(get.getComentarioNotaDeCredito());

            try {
                System.out.println(get.getTotalNCuota());
                double firstNumber = Double.valueOf(get.getTotalNCuota());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[12] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[13] = get.getProcentaje();

            String comnetario = get.getComnetario();
            System.out.println("---------------------------------------------> comnetario " + comnetario);
            if (comnetario != null) {
                fila[14] = comnetario;
            } else {
                fila[14] = "";
            }

            model.addRow(fila);
        }
        jTable3.setModel(model);

        trs = new TableRowSorter(model);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

//        trs.setComparator(3, new IntComparator());
//        trs.setComparator(4, new StringComparator());
//        trs.setComparator(5, new StringComparator());
//        trs.setComparator(6, new StringComparator());
        jTable3.setRowSorter(trs);
//                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
//                renderer.setHorizontalAlignment(JLabel.CENTER);
        Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

//        jTable2.getColumnModel().getColumn(0).setWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMinWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMaxWidth(0);
        jTable3.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(5).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(6).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(7).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(8).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
        jTable3.getColumnModel().getColumn(10).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(11).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(12).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(13).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(14).setCellRenderer(Renderer);
        jTable3.getColumnModel().getColumn(15).setCellRenderer(Renderer);

        jTable3.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);

        jTable3.getColumnModel().getColumn(1).setPreferredWidth(30);
        jTable3.getColumnModel().getColumn(2).setPreferredWidth(30);
        jTable3.getColumnModel().getColumn(3).setPreferredWidth(30);
        jTable3.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable3.getColumnModel().getColumn(11).setPreferredWidth(150);
        jTable3.getColumnModel().getColumn(14).setPreferredWidth(250);
        jTable3.getColumnModel().getColumn(15).setPreferredWidth(75);

        jTable3.setRowHeight(80);
        jTable3.setShowHorizontalLines(true);
        jTable3.setShowVerticalLines(true);

        jTable3.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable3));
        jTable3.getColumn("comentario").setCellEditor(new CustomEditor(jTable3));

        jTable3.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable3));
        jTable3.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable3));

        jTable3.getColumn("estado").setCellRenderer(new CustomRenderer1("OK?"));
        jTable3.getColumn("estado").setCellEditor(new CustomEditor1("OK?", jTable3));

        jTable3.getColumn("estado1").setCellRenderer(new CustomRenderer1("OC?"));
        jTable3.getColumn("estado1").setCellEditor(new CustomEditor1("OC?", jTable3));
        ////////////////////////////////////////////////////////////////////////

        model = new DefaultTableModel();
        model.addColumn("pkNumeroCuota");
        model.addColumn("numeroOrdenDoc");
        model.addColumn("guiaChilemat");
        model.addColumn("guiaProveedor");
        model.addColumn("local");
//        model.addColumn("numeroOrdenIngreso");
        model.addColumn("numeroGuia");
//        model.addColumn("fechaRecepcion");
        model.addColumn("sucursal");
        model.addColumn("proveedor");
        model.addColumn("fechaEmision");
        model.addColumn("fechaVencimiento");
//        model.addColumn("total");
        model.addColumn("saldo");
        model.addColumn("Nota de Credito");
        model.addColumn("totalNCuota");
        model.addColumn("procentaje");
        model.addColumn("comentario");
        model.addColumn("estado");
        model.addColumn("estado1");

        fila = new Object[17];

//                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
//                sdf.setLenient(false);
        /*    
                String numeroOrdenDoc;
                String guiaChilemat;
                String guiaProveedor;
                String local;

                String numeroOrdenIngreso;
                String numeroGuia;
                String fechaRecepcion;
                String sucursal;

                String proveedor;
                String fechaEmision;
                String total;
                String saldo;

                String totalNCuota;
                String procentaje;
         */
        for (int i = 0; i < arrOC.size(); ++i) {
            System.out.println("--------------------------------------------> " + i);
            Cruze get = arrOC.get(i);

            fila[0] = get.getPkNumeroCuota();
            fila[1] = get.getNumeroOrdenDoc();
            fila[2] = get.getGuiaChilemat();
            fila[3] = get.getGuiaProveedor();
            fila[4] = get.getLocal();

//            fila[5] = get.getNumeroOrdenIngreso();
            fila[5] = get.getNumeroGuia();
//            fila[6] = get.getFechaRecepcion();
            fila[6] = get.getSucursal();

            fila[7] = get.getProveedor();
            fila[8] = get.getFechaEmision();

            try {
                System.out.println("get.getFechaVencimiento() " + get.getFechaVencimiento());
                Date date = new SimpleDateFormat("dd/MM/yyyy").parse(get.getFechaVencimiento());
                fila[9] = date;
            } catch (Exception ex) {
                fila[9] = "";
            }

//            fila[10] = get.getTotal();
            try {
                System.out.println(get.getSaldo());
                double firstNumber = Double.valueOf(get.getSaldo());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[10] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[11] = String.valueOf(get.getComentarioNotaDeCredito());

            try {
                System.out.println(get.getTotalNCuota());
                double firstNumber = Double.valueOf(get.getTotalNCuota());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[12] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[13] = get.getProcentaje();

            String comnetario = get.getComnetario();
            System.out.println("---------------------------------------------> comnetario " + comnetario);
            if (comnetario != null) {
                fila[14] = comnetario;
            } else {
                fila[14] = "";
            }

            model.addRow(fila);
        }
        jTable5.setModel(model);

        trs = new TableRowSorter(model);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

//        trs.setComparator(3, new IntComparator());
//        trs.setComparator(4, new StringComparator());
//        trs.setComparator(5, new StringComparator());
//        trs.setComparator(6, new StringComparator());
        jTable5.setRowSorter(trs);
//                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
//                renderer.setHorizontalAlignment(JLabel.CENTER);
        Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

//        jTable2.getColumnModel().getColumn(0).setWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMinWidth(0);
//        jTable2.getColumnModel().getColumn(0).setMaxWidth(0);
        jTable5.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(5).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(6).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(7).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(8).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
        jTable5.getColumnModel().getColumn(10).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(11).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(12).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(13).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(14).setCellRenderer(Renderer);
        jTable5.getColumnModel().getColumn(15).setCellRenderer(Renderer);

        jTable5.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);

        jTable5.getColumnModel().getColumn(1).setPreferredWidth(30);
        jTable5.getColumnModel().getColumn(2).setPreferredWidth(30);
        jTable5.getColumnModel().getColumn(3).setPreferredWidth(30);
        jTable5.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable5.getColumnModel().getColumn(11).setPreferredWidth(150);
        jTable5.getColumnModel().getColumn(14).setPreferredWidth(250);
        jTable5.getColumnModel().getColumn(15).setPreferredWidth(75);

        jTable5.setRowHeight(80);
        jTable5.setShowHorizontalLines(true);
        jTable5.setShowVerticalLines(true);

        jTable5.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable5));
        jTable5.getColumn("comentario").setCellEditor(new CustomEditor(jTable5));

        jTable5.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable5));
        jTable5.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable5));

        jTable5.getColumn("estado").setCellRenderer(new CustomRenderer1("OK?"));
        jTable5.getColumn("estado").setCellEditor(new CustomEditor1("OK?", jTable5));

        jTable5.getColumn("estado1").setCellRenderer(new CustomRenderer1("MALA?"));
        jTable5.getColumn("estado1").setCellEditor(new CustomEditor1("MALA?", jTable5));
        ////////////////////////////////////////////////////////////////////////

        model = new DefaultTableModel();

        model.addColumn("pkNumeroCuota");
        model.addColumn("numeroOrdenDoc");
        model.addColumn("guiaChilemat");
        model.addColumn("guiaProveedor");
        model.addColumn("local");
//        model.addColumn("numeroOrdenIngreso");
        model.addColumn("numeroGuia");
//        model.addColumn("fechaRecepcion");
        model.addColumn("sucursal");
        model.addColumn("proveedor");
        model.addColumn("fechaEmision");
        model.addColumn("fechaVencimiento");
//        model.addColumn("total");
        model.addColumn("saldo");
        model.addColumn("Nota de Credito");
        model.addColumn("totalNCuota");
        model.addColumn("procentaje");
        model.addColumn("comentario");
        model.addColumn("estado");

        fila = new Object[16];

//                SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
//                sdf.setLenient(false);
        /*    
                String numeroOrdenDoc;
                String guiaChilemat;
                String guiaProveedor;
                String local;

                String numeroOrdenIngreso;
                String numeroGuia;
                String fechaRecepcion;
                String sucursal;

                String proveedor;
                String fechaEmision;
                String total;
                String saldo;

                String totalNCuota;
                String procentaje;
         */
        for (int i = 0; i < arrSinCruzar.size(); ++i) {
            System.out.println("--------------------------------------------> " + i);
            Cruze get = arrSinCruzar.get(i);

            fila[0] = get.getPkNumeroCuota();
            fila[1] = get.getNumeroOrdenDoc();
            fila[2] = get.getGuiaChilemat();
            fila[3] = get.getGuiaProveedor();
            fila[4] = get.getLocal();

//           fila[5] = get.getNumeroOrdenIngreso();
            fila[5] = get.getNumeroGuia();
//            fila[6] = get.getFechaRecepcion();
            fila[6] = get.getSucursal();

            fila[7] = get.getProveedor();
            fila[8] = get.getFechaEmision();

            try {
                System.out.println("get.getFechaVencimiento() " + get.getFechaVencimiento());
                Date date = new SimpleDateFormat("dd/MM/yyyy").parse(get.getFechaVencimiento());
                fila[9] = date;
            } catch (Exception ex) {
                fila[9] = "";
            }

            try {
                System.out.println(get.getSaldo());
                double firstNumber = Double.valueOf(get.getSaldo());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[10] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[11] = String.valueOf(get.getComentarioNotaDeCredito());

            try {
                System.out.println(get.getTotalNCuota());
                double firstNumber = Double.valueOf(get.getTotalNCuota());
                DecimalFormat decimalFormat = new DecimalFormat("#,##0");
                String firstNumberAsString = decimalFormat.format(firstNumber);
                fila[12] = firstNumberAsString;
            } catch (Exception ex) {
            }

            fila[13] = get.getProcentaje();

            String comnetario = get.getComnetario();
            System.out.println("---------------------------------------------> comnetario " + comnetario);
            if (comnetario != null) {
                fila[14] = comnetario;
            } else {
                fila[14] = "";
            }

            model.addRow(fila);
        }

        jTable6.setModel(model);

        trs = new TableRowSorter(model);

        trs.setComparator(8, new IntComparatorFecha0());
        trs.setComparator(9, new IntComparatorFecha1());

        trs.setComparator(10, new IntComparator());
        trs.setComparator(11, new IntComparator());
        trs.setComparator(12, new IntComparator());

//        trs.setComparator(12, new IntComparator());
//        trs.setComparator(2, new StringComparator2());
//        trs.setComparator(3, new IntComparator());
//        trs.setComparator(4, new StringComparator());
//        trs.setComparator(5, new StringComparator());
//        trs.setComparator(6, new StringComparator());
        jTable6.setRowSorter(trs);
//                DefaultTableCellRenderer renderer = (DefaultTableCellRenderer) jTable1.getTableHeader().getDefaultRenderer();
//                renderer.setHorizontalAlignment(JLabel.CENTER);

        centerRenderer2 = new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                try {
                    DateFormat dateFormat0 = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                    Date parse = dateFormat0.parse(value.toString());
                    DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                    String strDate = dateFormat.format(parse);
                    value = strDate;
                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                } catch (ParseException ex) {
                    System.out.println(ex);
                    return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                }
            }
        };
        centerRenderer2.setHorizontalAlignment(JLabel.CENTER);

        Renderer = new DefaultTableCellRenderer();
        Renderer.setHorizontalAlignment(JLabel.CENTER);

//        jTable1.getColumnModel().getColumn(0).setWidth(0);
//        jTable1.getColumnModel().getColumn(0).setMinWidth(0);
//        jTable1.getColumnModel().getColumn(0).setMaxWidth(0);
        jTable6.getColumnModel().getColumn(0).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(1).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(3).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(4).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(5).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(6).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(7).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(8).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(2).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(9).setCellRenderer(centerRenderer2);
        jTable6.getColumnModel().getColumn(10).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(11).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(12).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(13).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(14).setCellRenderer(Renderer);
        jTable6.getColumnModel().getColumn(15).setCellRenderer(Renderer);

        jTable6.setAutoResizeMode(JTable.AUTO_RESIZE_ALL_COLUMNS);

        jTable6.getColumnModel().getColumn(1).setPreferredWidth(30);
        jTable6.getColumnModel().getColumn(2).setPreferredWidth(30);
        jTable6.getColumnModel().getColumn(3).setPreferredWidth(30);
        jTable6.getColumnModel().getColumn(4).setPreferredWidth(5);
        jTable6.getColumnModel().getColumn(11).setPreferredWidth(150);
        jTable6.getColumnModel().getColumn(14).setPreferredWidth(250);
        jTable6.getColumnModel().getColumn(15).setPreferredWidth(75);

        jTable6.setRowHeight(80);
        jTable6.setShowHorizontalLines(true);
        jTable6.setShowVerticalLines(true);

        jTable6.getColumn("comentario").setCellRenderer(new CustomRenderer(jTable1));
        jTable6.getColumn("comentario").setCellEditor(new CustomEditor(jTable1));

        jTable6.getColumn("Nota de Credito").setCellRenderer(new CustomRenderer(jTable1));
        jTable6.getColumn("Nota de Credito").setCellEditor(new CustomEditor(jTable1));

        jTable6.getColumn("estado").setCellRenderer(new CustomRenderer1("REVISAR?"));
        jTable6.getColumn("estado").setCellEditor(new CustomEditor1("REVISAR?", jTable1));
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    public static javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton8;
    public static javax.swing.JComboBox<String> jComboBox1;
    public static javax.swing.JComboBox<String> jComboBox2;
    public static javax.swing.JComboBox<String> jComboBox3;
    private javax.swing.JLabel jLabel1;
    public static javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    public static javax.swing.JLabel jLabel12;
    public static javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel34;
    public static javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    public static javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    public static javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuBar jMenuBar2;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JPanel jPanel1;
    public static javax.swing.JPanel jPanel19;
    private javax.swing.JPanel jPanel2;
    public static javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    public static javax.swing.JPanel jPanel6;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JScrollPane jScrollPane5;
    private javax.swing.JScrollPane jScrollPane6;
    public static javax.swing.JTabbedPane jTabbedPane1;
    public static javax.swing.JTable jTable1;
    public static javax.swing.JTable jTable2;
    public static javax.swing.JTable jTable3;
    public static javax.swing.JTable jTable4;
    public static javax.swing.JTable jTable5;
    public static javax.swing.JTable jTable6;
    // End of variables declaration//GEN-END:variables
    public static String v = "";
    public static boolean opcion = true;
//    public static VentanaCargar cg;
}

class Editor_CheckBox extends AbstractCellEditor implements TableCellEditor, ActionListener {

    private final JCheckBox jCheckBox;

    public Editor_CheckBox() {
//        System.out.println("Editor" + " -> " + "public Editor_CheckBox()");
        jCheckBox = new JCheckBox();
        jCheckBox.addActionListener(this);
    }

    @Override
    public Object getCellEditorValue() {
//        System.out.println("Editor" + " -> " + "public Object getCellEditorValue()");
        return jCheckBox.isSelected();
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
//        System.out.println("Editor" + " -> " + " public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column)");
        jCheckBox.setHorizontalAlignment(CENTER);
        boolean b = (Boolean) value;
        jCheckBox.setSelected(b);
        return jCheckBox;
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        int selectedRow = jTable4.getSelectedRow();
        String fechas = jTable4.getValueAt(selectedRow, 0).toString();
        String comentario = jTable4.getValueAt(selectedRow, 7).toString();
        boolean selected = jCheckBox.isSelected();
        System.out.println("fechas " + fechas);
        System.out.println("selected " + selected);
        System.out.println("comentario " + comentario);
        int num = 999999999;

        if (selected) {
            jTable4.setValueAt(true, selectedRow, 6);
            num = 0;
        } else {
            jTable4.setValueAt(false, selectedRow, 6);
            num = 1;
        }

        System.out.println("num " + num);

        Totales totales = new Totales();
        totales.setFechas(fechas);
        totales.setSelected(num);
        totales.setComentario(comentario);

        try {
            TotalesDAO.registraTotales(totales, "ingresos");
        } catch (Exception ex) {
            System.out.println("ex " + ex);
            try {
                TotalesDAO.actualizaSelected(totales, "ingresos");
            } catch (IOException ex1) {
                System.out.println("ex1 " + ex1);
                Logger.getLogger(Editor_CheckBox.class.getName()).log(Level.SEVERE, null, ex1);
            } catch (SQLException ex1) {
                Exceptions.printStackTrace(ex1);
            }
        }
        suma();
        NewJFrame.selectedTabla4();
        NewJFrame.corrigueSelected();
    }
}

////////////////////////////////////////////////////////////////////////////////
class Render_CheckBox extends JCheckBox implements TableCellRenderer {

    private static final long serialVersionUID = 1L;
    private final JCheckBox jCheckBox = new JCheckBox();

    public Render_CheckBox() {
//        System.out.println("Render" + " -> " + "public Render_CheckBox()");
        setOpaque(true);
    }

    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
//        System.out.println("Render" + " -> " + "public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column)");
        jCheckBox.setHorizontalAlignment(CENTER);
        boolean b = (Boolean) value;
        jCheckBox.setSelected(b);
        return jCheckBox;
    }
}
////////////////////////////////////////////////////////////////////////////////

class CustomRenderer implements TableCellRenderer {

    JPanel panel;
    JScrollPane scrollPane;
    JTextArea textArea;
    JButton b;

    public CustomRenderer(JTable jTable) {
        panel = new javax.swing.JPanel();
        b = new javax.swing.JButton();
        scrollPane = new javax.swing.JScrollPane();
        textArea = new javax.swing.JTextArea();

        b.setText("Guardar");

        textArea.setColumns(18);
        int columnCount = jTable.getColumnCount();
        if (columnCount == 4) {
            textArea.setRows(1);
        } else {
            textArea.setRows(3);
        }
        scrollPane.setViewportView(textArea);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(panel);
        panel.setLayout(jPanel1Layout);

        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(scrollPane, javax.swing.GroupLayout.DEFAULT_SIZE, 259, Short.MAX_VALUE)
                                .addGap(0, 0, 0)
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 113, Short.MAX_VALUE)
                                .addGap(0, 0, 0))
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(scrollPane, javax.swing.GroupLayout.DEFAULT_SIZE, 146, Short.MAX_VALUE)
                        .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
    }

    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
        textArea.setText((String) value);
        return panel;
    }
}

////////////////////////////////////////////////////////////////////////////////
class CustomEditor extends AbstractCellEditor implements TableCellEditor, ActionListener, FocusListener {

    JPanel panel;
    JTextArea textArea;
    JScrollPane scrollPane;
    JButton b;
    JTable jTableEditor;

    @Override
    public Object getCellEditorValue() {
        return textArea.getText();
    }

    public CustomEditor(JTable jTable) {
        this.jTableEditor = jTable;
        panel = new javax.swing.JPanel();
        b = new javax.swing.JButton();
        scrollPane = new javax.swing.JScrollPane();
        textArea = new javax.swing.JTextArea();

        b.setText("Guardar");
        b.addActionListener(this);
        textArea.addFocusListener(this);

        textArea.setColumns(18);

        int columnCount = jTable.getColumnCount();
        if (columnCount == 8) {
            textArea.setRows(1);
        } else {
            textArea.setRows(3);
        }
        scrollPane.setViewportView(textArea);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(panel);
        panel.setLayout(jPanel1Layout);

        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(scrollPane, javax.swing.GroupLayout.DEFAULT_SIZE, 259, Short.MAX_VALUE)
                                .addGap(0, 0, 0)
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 113, Short.MAX_VALUE)
                                .addGap(0, 0, 0))
        );

        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addComponent(scrollPane, javax.swing.GroupLayout.DEFAULT_SIZE, 146, Short.MAX_VALUE)
                        .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
    }

    public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
        textArea.setText((String) value);
        return panel;
    }

    @Override
    public void focusGained(FocusEvent e) {
        System.out.println("gain");
    }

    @Override
    public void focusLost(FocusEvent e) {
        System.out.println("lost");
        jTable1.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable2.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable3.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable4.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
    }

    public void actionPerformed(ActionEvent e) {
        if (e.getSource() instanceof JButton) {
            int selectedIndex = jTabbedPane1.getSelectedIndex();
            if (selectedIndex == 4) {
                int selectedRow = jTable4.getSelectedRow();
                String comentario = textArea.getText();
                String fechas = jTable4.getValueAt(selectedRow, 0).toString();
                String selected = jTable4.getValueAt(selectedRow, 6).toString();

                System.out.println("fechas " + fechas);
                System.out.println("selected " + selected);
                System.out.println("comentario " + comentario);

                Totales totales = new Totales();
                totales.setFechas(fechas);
                if ("true".equals(selected)) {
                    totales.setSelected(0);
                } else if ("false".equals(selected)) {
                    totales.setSelected(1);
                }
                totales.setComentario(comentario);
                try {
                    TotalesDAO.registraTotales(totales, "ingresos");
                } catch (Exception ex) {
                    try {
                        TotalesDAO.actualizaComentario(totales, "ingresos");
                    } catch (IOException ex1) {
                        Exceptions.printStackTrace(ex1);
                    } catch (SQLException ex1) {
                        Exceptions.printStackTrace(ex1);
                    }
                }
            } else {
                int selectedColumnx = jTableEditor.getSelectedColumn();
                if (selectedColumnx == 11) {
                    try {
                        String text = textArea.getText();
                        System.out.println("text " + text);
                        int selectedRow = jTableEditor.getSelectedRow();
                        int selectedColumn = jTableEditor.getSelectedColumn();
                        System.out.println(selectedRow + " " + selectedColumn);
                        String toString = jTableEditor.getValueAt(selectedRow, 11).toString();
                        System.out.println("toString " + toString);
                        String toString1 = jTableEditor.getValueAt(selectedRow, 0).toString();
                        System.out.println("toString1 " + toString1);

                        boolean actualizaComentarioNotaDeCredito = DocumentoCobranzaDAO.actualizaComentarioNotaDeCredito(text, toString1, "ingresos");
                        if (actualizaComentarioNotaDeCredito == true) {
                            jTableEditor.setValueAt(text, selectedRow, selectedColumn);
                        } else {
                            jTableEditor.setValueAt("0", selectedRow, selectedColumn);
                            textArea.setText("0");
                        }

                        suma();
                    } catch (Exception ex) {
                        System.out.println("ex " + ex);
                    }
                } else {
                    try {
                        String text = textArea.getText();
                        System.out.println("text " + text);
                        int selectedRow = jTableEditor.getSelectedRow();
                        int selectedColumn = jTableEditor.getSelectedColumn();
                        System.out.println(selectedRow + " " + selectedColumn);
                        /*
                        String toString = jTableEditor.getValueAt(selectedRow, 14).toString();
                        System.out.println("toString " + toString);
                         */
                        String toString1 = jTableEditor.getValueAt(selectedRow, 0).toString();
                        System.out.println("toString1 " + toString1);

                        DocumentoCobranzaDAO.actualizaComentarioDocumentoCobranza(text, toString1, "ingresos");
                        JOptionPane.showMessageDialog(null, "Comentario Guardado");
                    } catch (IOException ex) {
                        Logger.getLogger(CustomEditor.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (SQLException ex) {
                        Exceptions.printStackTrace(ex);
                    }
                }
            }
        }
    }
}

////////////////////////////////////////////////////////////////////////////////
class CustomRenderer1 implements TableCellRenderer {

    JPanel panel;
    JScrollPane scrollPane;
    JTextArea textArea;
    JButton b;

    public CustomRenderer1(String name) {
        panel = new javax.swing.JPanel();
        b = new javax.swing.JButton();
        scrollPane = new javax.swing.JScrollPane();
        textArea = new javax.swing.JTextArea();

        b.setText(name);

        textArea.setColumns(18);
        textArea.setRows(3);
        scrollPane.setViewportView(textArea);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(panel);
        panel.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 273, Short.MAX_VALUE)
                                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 165, Short.MAX_VALUE)
                                .addContainerGap())
        );
    }

    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
        textArea.setText((String) value);
        return panel;
    }
}
////////////////////////////////////////////////////////////////////////////////

class CustomEditor1 extends AbstractCellEditor implements TableCellEditor, ActionListener, FocusListener {

    JPanel panel;
    JTextArea textArea;
    JScrollPane scrollPane;
    JButton b;
    JTable JtableEditor;

    @Override
    public Object getCellEditorValue() {
        return textArea.getText();
    }

    public CustomEditor1(String nombre, JTable Jtable) {
        this.JtableEditor = Jtable;
        panel = new javax.swing.JPanel();
        b = new javax.swing.JButton();
        scrollPane = new javax.swing.JScrollPane();
        textArea = new javax.swing.JTextArea();

        b.setText(nombre);
        b.addActionListener(this);
        textArea.addFocusListener(this);

        textArea.setColumns(18);
        textArea.setRows(3);
        scrollPane.setViewportView(textArea);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(panel);
        panel.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 273, Short.MAX_VALUE)
                                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
                jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(jPanel1Layout.createSequentialGroup()
                                .addContainerGap()
                                .addComponent(b, javax.swing.GroupLayout.DEFAULT_SIZE, 165, Short.MAX_VALUE)
                                .addContainerGap())
        );
    }

    public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
        textArea.setText((String) value);
        return panel;
    }

    @Override
    public void focusGained(FocusEvent e) {
    }

    @Override
    public void focusLost(FocusEvent e) {
        System.out.println("lost");
        jTable1.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable2.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable3.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        jTable4.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
    }

    public void actionPerformed(ActionEvent e) {
        if (e.getSource() instanceof JButton) {
            try {
                b.requestFocus();
                int selectedIndex = jTabbedPane1.getSelectedIndex();

                NewJFrame.filtro = jButton2.getText();

                switch (selectedIndex) {
                    case 0: {
                        int selectedRow0 = jTable1.getSelectedRow();
                        int selectedRow;
                        try {
                            selectedRow = jTable1.getRowSorter().convertRowIndexToModel(selectedRow0);
                        } catch (Exception ex) {
                            selectedRow = selectedRow0;
                        }
                        System.out.println("1");
                        System.out.println("selectedRow " + selectedRow);
                        String pk = jTable1.getModel().getValueAt(selectedRow, 0).toString();
                        System.out.println("toString1 " + pk);
                        DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("MALO", pk, "ingresos");
                        String toString0;
                        String toString1;
                        String toString2;
                        String toString3;
                        String toString4;
                        String toString5;
                        String toString6;
                        String toString7;
                        String toString8;
                        String toString9;
                        String toString10;
                        String toString11;
                        String toString12;
                        String toString13;
                        String toString14;
                        try {
                            toString0 = jTable1.getModel().getValueAt(selectedRow, 0).toString();
                        } catch (Exception ex) {
                            toString0 = "";
                        }
                        try {
                            toString1 = jTable1.getModel().getValueAt(selectedRow, 1).toString();
                        } catch (Exception ex) {
                            toString1 = "";
                        }
                        try {
                            toString2 = jTable1.getModel().getValueAt(selectedRow, 2).toString();
                        } catch (Exception ex) {
                            toString2 = "";
                        }
                        try {
                            toString3 = jTable1.getModel().getValueAt(selectedRow, 3).toString();
                        } catch (Exception ex) {
                            toString3 = "";
                        }
                        try {
                            toString4 = jTable1.getModel().getValueAt(selectedRow, 4).toString();
                        } catch (Exception ex) {
                            toString4 = "";
                        }
                        try {
                            toString5 = jTable1.getModel().getValueAt(selectedRow, 5).toString();
                        } catch (Exception ex) {
                            toString5 = "";
                        }
                        try {
                            toString6 = jTable1.getModel().getValueAt(selectedRow, 6).toString();
                        } catch (Exception ex) {
                            toString6 = "";
                        }
                        try {
                            toString7 = jTable1.getModel().getValueAt(selectedRow, 7).toString();
                        } catch (Exception ex) {
                            toString7 = "";
                        }
                        try {
                            toString8 = jTable1.getModel().getValueAt(selectedRow, 8).toString();
                        } catch (Exception ex) {
                            toString8 = "";
                        }
                        try {
                            toString9 = jTable1.getModel().getValueAt(selectedRow, 9).toString();
                        } catch (Exception ex) {
                            toString9 = "";
                        }
                        try {
                            toString10 = jTable1.getModel().getValueAt(selectedRow, 10).toString();
                        } catch (Exception ex) {
                            toString10 = "";
                        }
                        try {
                            toString11 = jTable1.getModel().getValueAt(selectedRow, 11).toString();
                        } catch (Exception ex) {
                            toString11 = "";
                        }
                        try {
                            toString12 = jTable1.getModel().getValueAt(selectedRow, 12).toString();
                        } catch (Exception ex) {
                            toString12 = "";
                        }
                        try {
                            toString13 = jTable1.getModel().getValueAt(selectedRow, 13).toString();
                        } catch (Exception ex) {
                            toString13 = "";
                        }
                        try {
                            toString14 = jTable1.getModel().getValueAt(selectedRow, 14).toString();
                        } catch (Exception ex) {
                            toString14 = "";
                        }

                        DefaultTableModel model1 = (DefaultTableModel) jTable1.getModel();
                        DefaultTableModel model2 = (DefaultTableModel) jTable2.getModel();

                        Object[] fila = new Object[15];

                        fila[0] = toString0;
                        fila[1] = toString1;
                        fila[2] = toString2;
                        fila[3] = toString3;
                        fila[4] = toString4;
                        fila[5] = toString5;
                        fila[6] = toString6;
                        fila[7] = toString7;
                        fila[8] = toString8;

                        try {
                            DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                            Date parse = dateFormat.parse(toString9);

                            fila[9] = parse;
                        } catch (Exception ex) {
                            fila[9] = "";
                        }
                        System.out.println("-----------> toString10 " + toString10);
                        fila[10] = toString10;
                        fila[11] = toString11;
                        fila[12] = toString12;
                        fila[13] = toString13;
                        fila[14] = toString14;
                        int largo = jTable2.getRowCount();
                        if (selectedRow <= largo) {
                            model2.insertRow(selectedRow, fila);
                            model1.removeRow(selectedRow);
                        } else {
                            model2.addRow(fila);
                            model1.removeRow(selectedRow);
                        }
                        jTable1.repaint();
                        jTable2.repaint();
                        break;
                    }
                    case 1: {
                        int selectedColumn = jTable2.getSelectedColumn();
                        System.out.println("selectedColumn " + selectedColumn);

                        switch (selectedColumn) {
                            case 15: {
                                int selectedRow0 = jTable2.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable2.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }
                                System.out.println("2");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                String pre1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                System.out.println("toString1 " + pre);
                                System.out.println("pre1 " + pre1);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("BUENO", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;
                                try {
                                    toString0 = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable2.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable2.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable2.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable2.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable2.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable2.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable2.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable2.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable2.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable2.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable2.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable2.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable2.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable1.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable2.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }
                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;
                                int largo = jTable1.getRowCount();
                                if (selectedRow <= largo) {
                                    model1.insertRow(selectedRow, fila);
                                    model2.removeRow(selectedRow);
                                } else {
                                    model1.addRow(fila);
                                    model2.removeRow(selectedRow);
                                }
                                jTable1.repaint();
                                jTable2.repaint();
                                break;
                            }
                            case 16: {
                                int selectedRow0 = jTable2.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable2.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }
                                System.out.println("2");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                String pre1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                System.out.println("toString1 " + pre);
                                System.out.println("pre1 " + pre1);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("MUYMALO", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;

                                try {
                                    toString0 = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable2.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable2.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable2.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable2.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable2.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable2.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable2.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable2.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable2.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable2.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable2.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable2.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable2.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable3.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable2.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }

                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;

                                int largo = jTable3.getRowCount();
                                if (selectedRow <= largo) {
                                    model1.insertRow(selectedRow, fila);
                                    model2.removeRow(selectedRow);
                                } else {
                                    model1.addRow(fila);
                                    model2.removeRow(selectedRow);
                                }

                                jTable3.repaint();
                                jTable2.repaint();
                                break;
                            }
                            case 17: {
                                int selectedRow0 = jTable2.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable2.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }
                                System.out.println("2");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                String pre1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                System.out.println("toString1 " + pre);
                                System.out.println("pre1 " + pre1);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("OC", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;
                                try {
                                    toString0 = jTable2.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable2.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable2.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable2.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable2.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable2.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable2.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable2.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable2.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable2.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable2.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable2.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable2.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable2.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable2.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable5.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable2.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }
                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;
                                int largo = jTable5.getRowCount();
                                if (selectedRow <= largo) {
                                    model1.insertRow(selectedRow, fila);
                                    model2.removeRow(selectedRow);
                                } else {
                                    model1.addRow(fila);
                                    model2.removeRow(selectedRow);
                                }
                                jTable5.repaint();
                                jTable2.repaint();
                                break;
                            }
                            default:
                                break;
                        }
                        break;
                    }
                    case 2: {
                        int selectedColumn = jTable3.getSelectedColumn();
                        switch (selectedColumn) {
                            case 15: {
                                int selectedRow0 = jTable3.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable3.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }

                                System.out.println("1");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable3.getModel().getValueAt(selectedRow, 0).toString();
                                System.out.println("toString1 " + pre);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("BUENO", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;
                                try {
                                    toString0 = jTable3.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable3.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable3.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable3.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable3.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable3.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable3.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable3.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable3.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable3.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable3.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable3.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable3.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable3.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable3.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable3.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable1.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }
                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;
                                int largo = jTable1.getRowCount();
                                if (selectedRow <= largo) {
                                    model2.insertRow(selectedRow, fila);
                                    model1.removeRow(selectedRow);
                                } else {
                                    model2.addRow(fila);
                                    model1.removeRow(selectedRow);
                                }
                                jTable3.repaint();
                                jTable1.repaint();
                                break;
                            }
                            case 16: {
                                int selectedRow0 = jTable3.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable3.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }

                                System.out.println("1");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable3.getModel().getValueAt(selectedRow, 0).toString();
                                System.out.println("toString1 " + pre);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("OC", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;
                                try {
                                    toString0 = jTable3.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable3.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable3.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable3.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable3.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable3.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable3.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable3.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable3.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable3.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable3.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable3.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable3.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable3.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable3.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable3.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable5.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }
                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;
                                int largo = jTable5.getRowCount();
                                if (selectedRow <= largo) {
                                    model2.insertRow(selectedRow, fila);
                                    model1.removeRow(selectedRow);
                                } else {
                                    model2.addRow(fila);
                                    model1.removeRow(selectedRow);
                                }
                                jTable3.repaint();
                                jTable5.repaint();
                                break;
                            }
                            default:
                                break;
                        }
                        break;
                    }
                    case 3: {
                        int selectedColumn = jTable5.getSelectedColumn();
                        System.out.println("selectedColumn " + selectedColumn);

                        switch (selectedColumn) {
                            case 15: {
                                int selectedRow0 = jTable5.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable5.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }
                                System.out.println("2");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable5.getModel().getValueAt(selectedRow, 0).toString();
                                String pre1 = jTable5.getModel().getValueAt(selectedRow, 1).toString();
                                System.out.println("toString1 " + pre);
                                System.out.println("pre1 " + pre1);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("BUENO", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;
                                try {
                                    toString0 = jTable5.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable5.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable5.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable5.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable5.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable5.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable5.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable5.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable5.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable5.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable5.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable5.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable5.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable5.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable5.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable1.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable5.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }
                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;
                                int largo = jTable1.getRowCount();
                                if (selectedRow <= largo) {
                                    model1.insertRow(selectedRow, fila);
                                    model2.removeRow(selectedRow);
                                } else {
                                    model1.addRow(fila);
                                    model2.removeRow(selectedRow);
                                }
                                jTable1.repaint();
                                jTable5.repaint();
                                break;
                            }
                            case 16: {
                                int selectedRow0 = jTable5.getSelectedRow();
                                int selectedRow;
                                try {
                                    selectedRow = jTable5.getRowSorter().convertRowIndexToModel(selectedRow0);
                                } catch (Exception ex) {
                                    selectedRow = selectedRow0;
                                }

                                System.out.println("2");
                                System.out.println("selectedRow " + selectedRow);
                                String pre = jTable5.getModel().getValueAt(selectedRow, 0).toString();
                                String pre1 = jTable5.getModel().getValueAt(selectedRow, 1).toString();
                                System.out.println("toString1 " + pre);
                                System.out.println("pre1 " + pre1);
                                DocumentoCobranzaDAO.actualizaEstadoDocumentoCobranza("MUYMALO", pre, "ingresos");
                                String toString0;
                                String toString1;
                                String toString2;
                                String toString3;
                                String toString4;
                                String toString5;
                                String toString6;
                                String toString7;
                                String toString8;
                                String toString9;
                                String toString10;
                                String toString11;
                                String toString12;
                                String toString13;
                                String toString14;

                                try {
                                    toString0 = jTable5.getModel().getValueAt(selectedRow, 0).toString();
                                } catch (Exception ex) {
                                    toString0 = "";
                                }
                                try {
                                    toString1 = jTable5.getModel().getValueAt(selectedRow, 1).toString();
                                } catch (Exception ex) {
                                    toString1 = "";
                                }
                                try {
                                    toString2 = jTable5.getModel().getValueAt(selectedRow, 2).toString();
                                } catch (Exception ex) {
                                    toString2 = "";
                                }
                                try {
                                    toString3 = jTable5.getModel().getValueAt(selectedRow, 3).toString();
                                } catch (Exception ex) {
                                    toString3 = "";
                                }
                                try {
                                    toString4 = jTable5.getModel().getValueAt(selectedRow, 4).toString();
                                } catch (Exception ex) {
                                    toString4 = "";
                                }
                                try {
                                    toString5 = jTable2.getModel().getValueAt(selectedRow, 5).toString();
                                } catch (Exception ex) {
                                    toString5 = "";
                                }
                                try {
                                    toString6 = jTable5.getModel().getValueAt(selectedRow, 6).toString();
                                } catch (Exception ex) {
                                    toString6 = "";
                                }
                                try {
                                    toString7 = jTable5.getModel().getValueAt(selectedRow, 7).toString();
                                } catch (Exception ex) {
                                    toString7 = "";
                                }
                                try {
                                    toString8 = jTable5.getModel().getValueAt(selectedRow, 8).toString();
                                } catch (Exception ex) {
                                    toString8 = "";
                                }
                                try {
                                    toString9 = jTable5.getModel().getValueAt(selectedRow, 9).toString();
                                } catch (Exception ex) {
                                    toString9 = "";
                                }
                                try {
                                    toString10 = jTable5.getModel().getValueAt(selectedRow, 10).toString();
                                } catch (Exception ex) {
                                    toString10 = "";
                                }
                                try {
                                    toString11 = jTable5.getModel().getValueAt(selectedRow, 11).toString();
                                } catch (Exception ex) {
                                    toString11 = "";
                                }
                                try {
                                    toString12 = jTable5.getModel().getValueAt(selectedRow, 12).toString();
                                } catch (Exception ex) {
                                    toString12 = "";
                                }
                                try {
                                    toString13 = jTable5.getModel().getValueAt(selectedRow, 13).toString();
                                } catch (Exception ex) {
                                    toString13 = "";
                                }
                                try {
                                    toString14 = jTable5.getModel().getValueAt(selectedRow, 14).toString();
                                } catch (Exception ex) {
                                    toString14 = "";
                                }
                                DefaultTableModel model1 = (DefaultTableModel) jTable3.getModel();
                                DefaultTableModel model2 = (DefaultTableModel) jTable5.getModel();
                                Object[] fila = new Object[15];
                                fila[0] = toString0;
                                fila[1] = toString1;
                                fila[2] = toString2;
                                fila[3] = toString3;
                                fila[4] = toString4;
                                fila[5] = toString5;
                                fila[6] = toString6;
                                fila[7] = toString7;
                                fila[8] = toString8;

                                try {
                                    DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                                    Date parse = dateFormat.parse(toString9);

                                    fila[9] = parse;
                                } catch (Exception ex) {
                                    fila[9] = "";
                                }

                                fila[10] = toString10;
                                fila[11] = toString11;
                                fila[12] = toString12;
                                fila[13] = toString13;
                                fila[14] = toString14;

                                int largo = jTable3.getRowCount();
                                if (selectedRow <= largo) {
                                    model1.insertRow(selectedRow, fila);
                                    model2.removeRow(selectedRow);
                                } else {
                                    model1.addRow(fila);
                                    model2.removeRow(selectedRow);
                                }

                                jTable3.repaint();
                                jTable5.repaint();
                                break;
                            }
                            default:
                                break;
                        }
                    }
                    default:
                        break;
                }

                NewJFrame.llenaTabla4();
                jTable4.repaint();
            } catch (IOException ex) {
                Logger.getLogger(CustomEditor.class.getName()).log(Level.SEVERE, null, ex);
            } catch (ParseException ex) {
                Logger.getLogger(CustomEditor1.class.getName()).log(Level.SEVERE, null, ex);
            } catch (SQLException ex) {
                Exceptions.printStackTrace(ex);
            }
            NewJFrame.total();

            filterheader.setTable(null);
            filterheader1.setTable(null);
            filterheader2.setTable(null);
            filterheader4.setTable(null);

            NewJFrame.filtrarSemana2();

            suma();

            NewJFrame.selectedTabla4();
            NewJFrame.corrigueSelected();

            TableModel modelx1 = jTable1.getModel();
            TableRowSorter trs = new TableRowSorter(modelx1);

            trs.setComparator(8, new IntComparatorFecha0());
            trs.setComparator(9, new IntComparatorFecha1());

            trs.setComparator(10, new IntComparator());
            trs.setComparator(11, new IntComparator());
            trs.setComparator(12, new IntComparator());

            jTable1.setRowSorter(trs);

            TableModel modelx2 = jTable2.getModel();
            trs = new TableRowSorter(modelx2);

            trs.setComparator(8, new IntComparatorFecha0());
            trs.setComparator(9, new IntComparatorFecha1());

            trs.setComparator(10, new IntComparator());
            trs.setComparator(11, new IntComparator());
            trs.setComparator(12, new IntComparator());

            jTable2.setRowSorter(trs);

            TableModel modelx3 = jTable3.getModel();
            trs = new TableRowSorter(modelx3);

            trs.setComparator(8, new IntComparatorFecha0());
            trs.setComparator(9, new IntComparatorFecha1());

            trs.setComparator(10, new IntComparator());
            trs.setComparator(11, new IntComparator());
            trs.setComparator(12, new IntComparator());

            jTable3.setRowSorter(trs);

            TableModel modelx4 = jTable5.getModel();
            trs = new TableRowSorter(modelx4);

            trs.setComparator(8, new IntComparatorFecha0());
            trs.setComparator(9, new IntComparatorFecha1());

            trs.setComparator(10, new IntComparator());
            trs.setComparator(11, new IntComparator());
            trs.setComparator(12, new IntComparator());

            jTable5.setRowSorter(trs);

            filterheader.setTable(jTable1);
            filterheader1.setTable(jTable2);
            filterheader2.setTable(jTable3);
            filterheader4.setTable(jTable5);

            NewJFrame.filtrarSemana2();
        }
    }
}

class IntComparator implements Comparator {

    public int compare(Object obj1, Object obj2) {
        int p1 = Integer.valueOf(obj1.toString().replace(".", ""));
        int p2 = Integer.valueOf(obj2.toString().replace(".", ""));
        if (p1 > p2) {
            return 1;
        } else if (p1 < p2) {
            return -1;
        } else {
            return 0;
        }
    }
}

class IntComparatorFecha0 implements Comparator {

    public int compare(Object o1, Object o2) {
        String valueOf = String.valueOf(o1);
        String[] split = valueOf.split("/");
        String name = split[2] + split[1] + split[0];

        String valueOf1 = String.valueOf(o2);
        String[] split1 = valueOf1.split("/");
        String name1 = split1[2] + split1[1] + split1[0];

        Integer int1 = Integer.valueOf(name);
        Integer int2 = Integer.valueOf(name1);
        return int1.compareTo(int2);
    }

    public boolean equals(Object o2) {
        return o2.equals(this);
    }
}

class IntComparatorFecha1 implements Comparator {

    public int compare(Object o1, Object o2) {
        String valueOf = String.valueOf(o1);

        DateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        Date parse = null;
        try {
            parse = dateFormat.parse(valueOf);
        } catch (ParseException ex) {
            Exceptions.printStackTrace(ex);
            System.out.println("ex " + ex);
        }
        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        String strDate = formatter.format(parse);

        String[] split = strDate.split("/");
        String name = split[2] + split[1] + split[0];

        String valueOf1 = String.valueOf(o2);

        DateFormat dateFormat1 = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        Date parse1 = null;
        try {
            parse1 = dateFormat1.parse(valueOf1);
        } catch (ParseException ex) {
            Exceptions.printStackTrace(ex);
            System.out.println("ex " + ex);
        }
        SimpleDateFormat formatter1 = new SimpleDateFormat("dd/MM/yyyy");
        String strDate1 = formatter1.format(parse1);

        String[] split1 = strDate1.split("/");
        String name1 = split1[2] + split1[1] + split1[0];

        Integer int1 = Integer.valueOf(name);
        Integer int2 = Integer.valueOf(name1);
        return int1.compareTo(int2);
    }

    public boolean equals(Object o2) {
        return o2.equals(this);
    }
}
