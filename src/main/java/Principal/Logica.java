package Principal;

import Clases.Cruze;
import Clases.DocumentoCobranza;
import Clases.Ingreso;
import DAO.DocumentoCobranzaDAO;
import static DAO.DocumentoCobranzaDAO.consultaDocumentoCobranza2;
import DAO.IngresoDAO;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openide.util.Exceptions;

public class Logica {

    public static ArrayList<Cruze> arrCruzeOK = new ArrayList<>();
    public static ArrayList<Cruze> arrCruzeBAD = new ArrayList<>();
    public static ArrayList<Cruze> arrCruzeMuyMALA = new ArrayList<>();
    public static ArrayList<Cruze> arrOC = new ArrayList<>();
    public static ArrayList<Cruze> arrSinCruzar = new ArrayList<>();

    public static void main(String[] args) throws IOException, SQLException, ParseException {
//        leerExcel();
        Thread thread = new Thread(() -> {
            try {                
//                VentanaCargar vc = new VentanaCargar();
//                ImageIcon icon = new ImageIcon(System.getProperty("user.dir") + "\\Iconos\\loading.gif");
//                vc.jLabel1.setIcon(icon);
//                vc.setLocationRelativeTo(null);
//                vc.setVisible(true);
                manejo();
//                vc.dispose();
                NewJFrame nf = new NewJFrame();
                nf.setVisible(true);
            } catch (IOException ex) {
                Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
            } catch (SQLException ex) {
                Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        thread.start();
    }

    public static void leerExcel(File file) throws IOException {
        try {
            ArrayList<DocumentoCobranza> documentoCobranza = leerExcel2_2(file);
//            System.out.println("");
//            System.out.println(documentoCobranza.size());

//            System.exit(0);
            ArrayList<Integer> index = new ArrayList<>();
            AtomicInteger atomicInteger = new AtomicInteger(0);

            documentoCobranza.stream().forEach((DocumentoCobranza doc) -> {
                int numero = doc.getNumero();
                String tipo = doc.getTipo();
                int cuota = doc.getCuota();
                String sucursal = doc.getSucursal();
                String proveedor = doc.getProveedor();
                String fechaEmision = doc.getFechaEmision();
                String fechaVencimiento = doc.getFechaVencimiento();
                int montoCuota = doc.getMontoCuota();
                int saldo = doc.getSaldo();
//            int dias = doc.getDias();
                int numeroOrden = doc.getNumeroOrden();
                String guiaChilemat = doc.getGuiaChilemat();
                String guiaProveedor = doc.getGuiaProveedor();
                int numeroCuota = doc.getNumeroCuota();
                doc.setPkNumeroCuota(numero + "_" + cuota);
                String pkNumeroCuota = doc.getPkNumeroCuota();

                atomicInteger.getAndIncrement();
                String toString = atomicInteger.toString();
                int valueOf = Integer.valueOf(toString);

                if (numero == 0) {
                    index.add(valueOf);
//                } else {
//                    System.out.println("------------------  " + valueOf + "   ------------------");
//                    System.out.println("numero " + numero);
//                    System.out.println("tipo " + tipo);
//                    System.out.println("cuota " + cuota);
//                    System.out.println("sucursal " + sucursal);
//                    System.out.println("proveedor " + proveedor);
//                    System.out.println("fechaEmision " + fechaEmision);
//                    System.out.println("fechaVencimiento " + fechaVencimiento);
//                    System.out.println("montoCuota " + montoCuota);
//                    System.out.println("saldo " + saldo);
//                    System.out.println("numeroOrden " + numeroOrden);
//                    System.out.println("guiaChilemat " + guiaChilemat);
//                    System.out.println("guiaProveedor " + guiaProveedor);
//                    System.out.println("numeroCuota " + numeroCuota);
//                    System.out.println("pkNumeroCuota " + pkNumeroCuota);
//                    System.out.println("----------------------------------------------");
                }
            });

            index.stream().forEach((Integer i) -> {
                documentoCobranza.remove((int) i);
            });
            AtomicInteger cont = new AtomicInteger();

//            documentoCobranza.stream().forEach((DocumentoCobranza doc) -> {
            for (int i = 0; i < documentoCobranza.size(); i++) {
                DocumentoCobranza doc = documentoCobranza.get(i);

                System.out.println("documentoCobranza.size() " + documentoCobranza.size());
                System.out.println("cont " + cont);
//                NewJFrame.cg.jLabel3.setText("" + documentoCobranza.size());
//                NewJFrame.cg.jLabel4.setText("" + cont);
                try {
                    DocumentoCobranzaDAO.registraDocumentoCobranza(doc, "ingresos");
                } catch (IOException ex) {
                    try {
                        DocumentoCobranza consultaDocumentoCobranzaUnico = DocumentoCobranzaDAO.consultaDocumentoCobranzaUnico(doc.getPkNumeroCuota(), "ingresos");
                        String guiaChilematAntiguo = consultaDocumentoCobranzaUnico.getGuiaChilemat();
                        String guiaProveedorAntiguo = consultaDocumentoCobranzaUnico.getGuiaProveedor();
                        String guiaChilematNuevo = doc.getGuiaChilemat();
                        String guiaProveedorNuevo = doc.getGuiaProveedor();

                        if (guiaChilematAntiguo.equals("") && !guiaChilematNuevo.equals("")) {
                            DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaChilemat(doc, "ingresos");
                        }
                        if (!guiaChilematAntiguo.equals("") && !guiaChilematNuevo.equals("")) {
                            if (!guiaChilematAntiguo.equals(guiaChilematNuevo)) {
                                DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaChilemat(doc, "ingresos");
                            }
                        }

                        if (guiaProveedorAntiguo.equals("") && !guiaProveedorNuevo.equals("")) {
                            DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaProveedor(doc, "ingresos");
                        }
                        if (!guiaProveedorAntiguo.equals("") && !guiaProveedorNuevo.equals("")) {
                            if (!guiaProveedorAntiguo.equals(guiaProveedorNuevo)) {
                                DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaProveedor(doc, "ingresos");
                            }
                        }
                    } catch (IOException ex1) {
                        Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (SQLException ex1) {
                        Exceptions.printStackTrace(ex1);
                    }
                } catch (SQLException ex) {
//                    System.out.println("ex " + ex);
                    try {
                        DocumentoCobranza consultaDocumentoCobranzaUnico = DocumentoCobranzaDAO.consultaDocumentoCobranzaUnico(doc.getPkNumeroCuota(), "ingresos");
                        String guiaChilematAntiguo = consultaDocumentoCobranzaUnico.getGuiaChilemat();
                        String guiaProveedorAntiguo = consultaDocumentoCobranzaUnico.getGuiaProveedor();
                        String guiaChilematNuevo = doc.getGuiaChilemat();
                        String guiaProveedorNuevo = doc.getGuiaProveedor();

                        if (guiaChilematAntiguo.equals("") && !guiaChilematNuevo.equals("")) {
                            DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaChilemat(doc, "ingresos");
                        }
                        if (!guiaChilematAntiguo.equals("") && !guiaChilematNuevo.equals("")) {
                            if (!guiaChilematAntiguo.equals(guiaChilematNuevo)) {
                                DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaChilemat(doc, "ingresos");
                            }
                        }

                        if (guiaProveedorAntiguo.equals("") && !guiaProveedorNuevo.equals("")) {
                            DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaProveedor(doc, "ingresos");
                        }
                        if (!guiaProveedorAntiguo.equals("") && !guiaProveedorNuevo.equals("")) {
                            if (!guiaProveedorAntiguo.equals(guiaProveedorNuevo)) {
                                DocumentoCobranzaDAO.actualizaDocumentoCobranzaGuiaProveedor(doc, "ingresos");
                            }
                        }
                    } catch (IOException ex1) {
                        Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (SQLException ex1) {
                        Exceptions.printStackTrace(ex1);
                    }
                }
                cont.getAndIncrement();
//            });
            }
            JOptionPane.showMessageDialog(null, "Proceso Terminado!");

//            System.out.println("leerExcel.size() " + documentoCobranza.size());
        } catch (IOException ex) {
            Exceptions.printStackTrace(ex);
        }
    }

    public static void manejo() throws IOException, SQLException {
        arrCruzeOK.clear();
        arrCruzeBAD.clear();
        arrCruzeMuyMALA.clear();
        arrOC.clear();
        arrSinCruzar.clear();

        ArrayList<DocumentoCobranza> consultaDocumentoCobranza = DocumentoCobranzaDAO.consultaDocumentoCobranza("ingresos");
//        System.out.println(consultaDocumentoCobranza.size());

        ArrayList<DocumentoCobranza> consultaDocumentoCobranza2 = consultaDocumentoCobranza2("ingresos");

        consultaDocumentoCobranza2.stream().forEach(new Consumer<DocumentoCobranza>() {
            @Override
            public void accept(DocumentoCobranza doc) {
                /*  String numeroOrdenDoc;
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
                    String pkNumeroCuota;
                    String comnetario;
                    String fechaVencimiento;
                    int comentarioNotaDeCredito;
                 */
                Cruze cruze = new Cruze();
                cruze.setNumeroOrdenDoc("" + doc.getNumeroOrden());
                cruze.setGuiaChilemat("" + doc.getGuiaChilemat());
                cruze.setGuiaProveedor("" + doc.getGuiaProveedor());
                cruze.setLocal("");
                cruze.setNumeroOrdenIngreso("");
                cruze.setNumeroGuia("");
                cruze.setFechaRecepcion("");
                cruze.setSucursal(doc.getSucursal());
                cruze.setProveedor(doc.getProveedor());
                cruze.setFechaEmision(doc.getFechaEmision());
                cruze.setTotal("");
                cruze.setSaldo("" + doc.getSaldo());
                cruze.setTotalNCuota("");
                cruze.setPorcentaje("");
                cruze.setPorcentaje("");
                cruze.setPkNumeroCuota(doc.getPkNumeroCuota());
                cruze.setComnetario(doc.getComentario());
                cruze.setFechaVencimiento(doc.getFechaVencimiento());
                cruze.setComentarioNotaDeCredito(doc.getComentarioNotaDeCrefito());
                arrSinCruzar.add(cruze);
            }
        });

        ArrayList<Ingreso> consultaIngresoVA = IngresoDAO.consultaIngresoVA("ingresos");
//        System.out.println(consultaIngresoVA.size());

        ArrayList<Ingreso> consultaIngresoPB = IngresoDAO.consultaIngresoPB("ingresos");
//        System.out.println(consultaIngresoPB.size());

        ArrayList<Ingreso> consultaIngresoOL = IngresoDAO.consultaIngresoOL("ingresos");
//        System.out.println(consultaIngresoOL.size());

        AtomicInteger counter = new AtomicInteger(0);

        consultaDocumentoCobranza.stream().forEach(new Consumer<DocumentoCobranza>() {
            @Override
            public void accept(DocumentoCobranza doc) {
                AtomicInteger at = new AtomicInteger(0);

                Cruze cruze = new Cruze();
                int andIncrement = counter.getAndIncrement();

                String estado = doc.getEstado();
                String pkNumeroCuota = doc.getPkNumeroCuota();
                String comentario = doc.getComentario();
                int comentarioNotaDeCredito = doc.getComentarioNotaDeCrefito();
//                System.out.println("-----------------------------> comentario " + comentario);

                int numeroOrdenDocInt = doc.getNumeroOrden();
                String numeroOrdenDoc = String.valueOf(numeroOrdenDocInt);

                String guiaChilematInt = doc.getGuiaChilemat();
                String guiaChilemat = String.valueOf(guiaChilematInt);
                String guiaProveedorInt = doc.getGuiaProveedor();
                String guiaProveedor = String.valueOf(guiaProveedorInt);
                int cuota = doc.getCuota();

//                System.out.println("---VENCIMIENTO---" + andIncrement + "---");
//                System.out.println("DOC numeroOrdenDoc " + numeroOrdenDoc);
                cruze.setNumeroOrdenDoc(String.valueOf(numeroOrdenDoc));
//                System.out.println("DOC guiaChilemat " + guiaChilemat);
                cruze.setGuiaChilemat(String.valueOf(guiaChilemat));
//                System.out.println("DOC guiaProveedor " + guiaProveedor);
                cruze.setGuiaProveedor(String.valueOf(guiaProveedor));

                ArrayList<String> arrOrdenDeCompra = new ArrayList<>();
                ArrayList<String> arrNumeroGuia = new ArrayList<>();
                ArrayList<String> arrFechaRecepcion = new ArrayList<>();
                ArrayList<Double> arrTotalNCuota = new ArrayList<>();

                consultaIngresoVA.stream().forEach((Ingreso ingreso) -> {
                    String ordenDeCompra = ingreso.getOrdenDeCompra();
                    Integer numeroOrdenIngresoInt = Integer.valueOf(ordenDeCompra);
                    String numeroOrdenIngreso = String.valueOf(numeroOrdenIngresoInt);

                    String estadoFolio = ingreso.getEstadoFolio();

                    if (numeroOrdenDoc.equals(numeroOrdenIngreso) && !estadoFolio.equals("NO VIGENTE")) {
                        String numeroGuia = ingreso.getNumeroGuia();
                        int numeroGuiaNInt = Integer.valueOf(numeroGuia);
                        String fechaRecepcion = ingreso.getFechaRecepcion();
                        Double valueOf = Double.valueOf(fechaRecepcion);

                        String numeroGuiaN = String.valueOf(numeroGuiaNInt);

                        Date javaDate = DateUtil.getJavaDate((double) valueOf);

                        if (numeroOrdenDoc.contains(numeroGuiaN) || guiaChilemat.contains(numeroGuiaN) || guiaProveedor.contains(numeroGuiaN)) {
//                            System.out.println("---INGRESO--- VA");
                            cruze.setLocal("VA");
//                            System.out.println("------numeroOrdenIngreso " + ordenDeCompra);
                            arrOrdenDeCompra.add(ordenDeCompra);
                            cruze.setNumeroOrdenIngreso(String.valueOf(ordenDeCompra));
//                            System.out.println("------numeroGuia " + numeroGuia);
                            arrNumeroGuia.add(numeroGuia);
                            cruze.setNumeroGuia(numeroGuia);
//                            System.out.println("------fechaRecepcion " + new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            arrFechaRecepcion.add(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            cruze.setFechaRecepcion(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));

                            String sucursal = doc.getSucursal();
                            String proveedor = doc.getProveedor();
                            String fechaEmision = doc.getFechaEmision();
                            double saldo = doc.getSaldo();
                            String total = ingreso.getTotal();
                            int totalN = Integer.valueOf(total);
                            double totalNCuota = totalN / cuota;
                            arrTotalNCuota.add(totalNCuota);

//                            System.out.println("DOC sucursal: " + sucursal);
                            cruze.setSucursal(sucursal);
//                            System.out.println("DOC proveedor " + proveedor);
                            cruze.setProveedor(proveedor);
//                            System.out.println("DOC fechaEmision " + fechaEmision);
                            cruze.setFechaEmision(fechaEmision);
//                            System.out.println("INGRESO total " + total);
                            cruze.setTotal(total);
//                            System.out.println("DOC saldo " + saldo);
                            cruze.setSaldo(String.valueOf(saldo));
//                            System.out.println("totalNCuota " + totalNCuota);
                            cruze.setTotalNCuota(String.valueOf(totalNCuota));

                            String fechaVencimiento = doc.getFechaVencimiento();
                            cruze.setFechaVencimiento(fechaVencimiento);

                            double procentaje = (saldo - totalNCuota) / totalNCuota;
                            String sValue = (String) String.format("%.3f", procentaje);

//                            System.out.println("procentaje " + (sValue));
                            cruze.setPorcentaje(String.valueOf(sValue));

                            at.getAndIncrement();
                        }
//                        System.out.println("--------------------------------------------------");
                    }
                });

                consultaIngresoPB.stream().forEach((Ingreso ingreso) -> {
                    String ordenDeCompra = ingreso.getOrdenDeCompra();
                    Integer numeroOrdenIngresoInt = Integer.valueOf(ordenDeCompra);
                    String numeroOrdenIngreso = String.valueOf(numeroOrdenIngresoInt);

                    String estadoFolio = ingreso.getEstadoFolio();

                    if (numeroOrdenDoc.equals(numeroOrdenIngreso) && !estadoFolio.equals("NO VIGENTE")) {
                        String numeroGuia = ingreso.getNumeroGuia();
                        int numeroGuiaNInt = Integer.valueOf(numeroGuia);
                        String fechaRecepcion = ingreso.getFechaRecepcion();
                        Double valueOf = Double.valueOf(fechaRecepcion);

                        String numeroGuiaN = String.valueOf(numeroGuiaNInt);

                        Date javaDate = DateUtil.getJavaDate((double) valueOf);

                        if (numeroOrdenDoc.contains(numeroGuiaN) || guiaChilemat.contains(numeroGuiaN) || guiaProveedor.contains(numeroGuiaN)) {
//                            System.out.println("---INGRESO--- PB");
                            cruze.setLocal("PB");
//                            System.out.println("------numeroOrdenIngreso " + ordenDeCompra);
                            arrOrdenDeCompra.add(ordenDeCompra);
                            cruze.setNumeroOrdenIngreso(String.valueOf(ordenDeCompra));
//                            System.out.println("------numeroGuia " + numeroGuia);
                            arrNumeroGuia.add(numeroGuia);
                            cruze.setNumeroGuia(numeroGuia);
//                            System.out.println("------fechaRecepcion " + new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            arrFechaRecepcion.add(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            cruze.setFechaRecepcion(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));

                            String sucursal = doc.getSucursal();
                            String proveedor = doc.getProveedor();
                            String fechaEmision = doc.getFechaEmision();
                            double saldo = doc.getSaldo();
                            String total = ingreso.getTotal();
                            int totalN = Integer.valueOf(total);
                            double totalNCuota = totalN / cuota;
                            arrTotalNCuota.add(totalNCuota);

//                            System.out.println("DOC sucursal: " + sucursal);
                            cruze.setSucursal(sucursal);
//                            System.out.println("DOC proveedor " + proveedor);
                            cruze.setProveedor(proveedor);
//                            System.out.println("DOC fechaEmision " + fechaEmision);
                            cruze.setFechaEmision(fechaEmision);
//                            System.out.println("INGRESO total " + total);
                            cruze.setTotal(total);
//                            System.out.println("DOC saldo " + saldo);
                            cruze.setSaldo(String.valueOf(saldo));
//                            System.out.println("totalNCuota " + totalNCuota);
                            cruze.setTotalNCuota(String.valueOf(totalNCuota));

                            String fechaVencimiento = doc.getFechaVencimiento();
                            cruze.setFechaVencimiento(fechaVencimiento);

                            double procentaje = (saldo - totalNCuota) / totalNCuota;
                            String sValue = (String) String.format("%.3f", procentaje);

//                            System.out.println("procentaje " + (sValue));
                            cruze.setPorcentaje(String.valueOf(sValue));

                            at.getAndIncrement();
                        }
//                        System.out.println("--------------------------------------------------");
                    }
                });

                consultaIngresoOL.stream().forEach((Ingreso ingreso) -> {
                    String ordenDeCompra = ingreso.getOrdenDeCompra();
                    Integer numeroOrdenIngresoInt = Integer.valueOf(ordenDeCompra);
                    String numeroOrdenIngreso = String.valueOf(numeroOrdenIngresoInt);

                    String estadoFolio = ingreso.getEstadoFolio();

                    if (numeroOrdenDoc.equals(numeroOrdenIngreso) && !estadoFolio.equals("NO VIGENTE")) {
                        String numeroGuia = ingreso.getNumeroGuia();
                        int numeroGuiaNInt = Integer.valueOf(numeroGuia);
                        String numeroGuiaN = String.valueOf(numeroGuiaNInt);
                        String fechaRecepcion = ingreso.getFechaRecepcion();
                        Double valueOf = Double.valueOf(fechaRecepcion);

                        Date javaDate = DateUtil.getJavaDate((double) valueOf);

                        if (numeroOrdenDoc.contains(numeroGuiaN) || guiaChilemat.contains(numeroGuiaN) || guiaProveedor.contains(numeroGuiaN)) {
//                            System.out.println("---INGRESO--- OL");
                            cruze.setLocal("OL");
//                            System.out.println("------numeroOrdenIngreso " + ordenDeCompra);
                            arrOrdenDeCompra.add(ordenDeCompra);
                            cruze.setNumeroOrdenIngreso(String.valueOf(ordenDeCompra));
//                            System.out.println("------numeroGuia " + numeroGuia);
                            arrNumeroGuia.add(numeroGuia);
                            cruze.setNumeroGuia(numeroGuia);
//                            System.out.println("------fechaRecepcion " + new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            arrFechaRecepcion.add(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));
                            cruze.setFechaRecepcion(new SimpleDateFormat("dd/MM/yyyy").format(javaDate));

                            String sucursal = doc.getSucursal();
                            String proveedor = doc.getProveedor();
                            String fechaEmision = doc.getFechaEmision();
                            double saldo = doc.getSaldo();
                            String total = ingreso.getTotal();
                            int totalN = Integer.valueOf(total);
                            double totalNCuota = totalN / cuota;
                            arrTotalNCuota.add(totalNCuota);

//                            System.out.println("DOC sucursal: " + sucursal);
                            cruze.setSucursal(sucursal);
//                            System.out.println("DOC proveedor " + proveedor);
                            cruze.setProveedor(proveedor);
//                            System.out.println("DOC fechaEmision " + fechaEmision);
                            cruze.setFechaEmision(fechaEmision);
//                            System.out.println("INGRESO total " + total);
                            cruze.setTotal(total);
//                            System.out.println("DOC saldo " + saldo);
                            cruze.setSaldo(String.valueOf(saldo));
//                            System.out.println("totalNCuota " + totalNCuota);
                            cruze.setTotalNCuota(String.valueOf(totalNCuota));

                            String fechaVencimiento = doc.getFechaVencimiento();
                            cruze.setFechaVencimiento(fechaVencimiento);

                            double procentaje = (saldo - totalNCuota) / totalNCuota;
                            String sValue = (String) String.format("%.3f", procentaje);

//                            System.out.println("procentaje " + (sValue));
                            cruze.setPorcentaje(String.valueOf(sValue));

                            at.getAndIncrement();
                        }
//                        System.out.println("--------------------------------------------------");
                    }
                });

                try {
                    Double saldo = Double.valueOf(cruze.getSaldo());
                    Double totalNCuota = Double.valueOf(cruze.getTotalNCuota());
                    double procentaje = (saldo - totalNCuota) / totalNCuota;

//                    System.out.println("saldo " + saldo);
//                    System.out.println("totalNCuota " + totalNCuota);
//                    System.out.println("procentaje " + procentaje);
                    String sValue = (String) String.format("%.3f", procentaje);
                    cruze.setPorcentaje(String.valueOf(sValue));
                } catch (Exception ex) {
                }

                String sumaOrdenDeCompra = "";
                for (int i = 0; i < arrOrdenDeCompra.size(); i++) {
                    String get = arrOrdenDeCompra.get(i);
                    sumaOrdenDeCompra = sumaOrdenDeCompra + " - " + get;
                }
                cruze.setNumeroOrdenIngreso(sumaOrdenDeCompra.replaceFirst("-", ""));

                String sumaNumeroGuia = "";
                for (int i = 0; i < arrNumeroGuia.size(); i++) {
                    String get = arrNumeroGuia.get(i);
                    sumaNumeroGuia = sumaNumeroGuia + " - " + get;
                }
                cruze.setNumeroGuia(sumaNumeroGuia.replaceFirst("-", ""));

                String sumaFechaRecepcion = "";
                for (int i = 0; i < arrFechaRecepcion.size(); i++) {
                    String get = arrFechaRecepcion.get(i);
                    sumaFechaRecepcion = sumaFechaRecepcion + " - " + get;
                }
                cruze.setFechaRecepcion(sumaFechaRecepcion.replaceFirst("-", ""));

                double sumaDouble = 0.0;
                for (int i = 0; i < arrTotalNCuota.size(); i++) {
                    double get = arrTotalNCuota.get(i);
                    sumaDouble = sumaDouble + get;
                }

                cruze.setTotalNCuota(String.valueOf(sumaDouble));
                String porcentajeString = String.valueOf(cruze.getProcentaje());

                cruze.setPkNumeroCuota(pkNumeroCuota);
                cruze.setComnetario(comentario);
                cruze.setComentarioNotaDeCredito(comentarioNotaDeCredito);

                int get = at.get();
                if (get == 0) {
                    cruze.setFechaVencimiento(String.valueOf(doc.getFechaVencimiento()));
                    cruze.setSaldo(String.valueOf(doc.getSaldo()));
                    cruze.setProveedor(String.valueOf(doc.getProveedor()));
                    cruze.setFechaEmision(String.valueOf(doc.getFechaEmision()));
                    cruze.setSucursal(String.valueOf(doc.getSucursal()));
                }

//                System.out.println(estado);
                if (estado == null) {
                    estado = "";
                }
                if (estado.equals("MALO")) {
                    arrCruzeBAD.add(cruze);
                } else if (estado.equals("BUENO")) {
                    arrCruzeOK.add(cruze);
                } else if (estado.equals("MUYMALO")) {
                    arrCruzeMuyMALA.add(cruze);
                } else if (estado.equals("OC")) {
                    arrOC.add(cruze);
                } else if (porcentajeString.equals("0,002")) {
                    arrCruzeOK.add(cruze);
                } else {
                    arrCruzeBAD.add(cruze);
                }

//                System.out.println("''''''''''''''''''''''''''''''''''''''''''''''''''''''''''");
            }
        });
    }

    public static ArrayList<DocumentoCobranza> leerExcel2_2(File file) throws IOException {
        ArrayList<DocumentoCobranza> arrIngreso = new ArrayList<>();

        FileInputStream ExcelFileToRead = new FileInputStream(file);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
        int g = 0;
        Row row = sheet.getRow(g);
        String name = "";

        while (!name.equals("NUMERO")) {
            try {
                Cell cel = row.getCell(0);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        name = String.valueOf(cel.getNumericCellValue());
//                        System.out.println(name);
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        name = cel.getStringCellValue();
//                        System.out.println(name);
                    }
                } else {
                    try {
                        name = cel.getStringCellValue();
//                        System.out.println(name);
                    } catch (Exception e) {
                        name = String.valueOf(cel.getNumericCellValue());
//                        System.out.println(name);
                    }
                }
                g++;
                row = sheet.getRow(g);
//                System.out.println("g " + g);
            } catch (Exception ex) {
                g++;
                row = sheet.getRow(g);
//                System.out.println("g " + g);
            }
        }

//        System.out.println("g " + g);
        XSSFRow row2 = sheet.getRow(g - 1);
        Cell cel = row2.getCell(9);
        if (cel.getCellType() == CellType.FORMULA) {
            if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                name = String.valueOf(cel.getNumericCellValue());
//                System.out.println(name);
            } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                name = cel.getStringCellValue();
//                System.out.println(name);
            }
        } else {
            try {
                name = cel.getStringCellValue();
//                System.out.println(name);
            } catch (Exception e) {
                name = String.valueOf(cel.getNumericCellValue());
//                System.out.println(name);
            }
        }

//        System.exit(0);
        if (!name.equals("DIAS")) {
            String valueOfx = "ABC";
            while (!"".equals(valueOfx)) {
                DocumentoCobranza ingreso = new DocumentoCobranza();
                cel = row.getCell(0);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumero((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumero(valueOf);
                    }
                }

                cel = row.getCell(1);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setTipo(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setTipo(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setTipo(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setTipo(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(2);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setCuota(valueOf);
                    }
                }

                cel = row.getCell(3);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setSucursal(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setSucursal(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setSucursal(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setSucursal(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(4);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setProveedor(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setProveedor(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setProveedor(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setProveedor(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(5);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setFechaEmision(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setFechaEmision(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setFechaEmision(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setFechaEmision(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(6);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setFechaVencimiento(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setFechaVencimiento(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setFechaVencimiento(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setFechaVencimiento(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(7);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setMontoCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setMontoCuota(valueOf);
                    }
                }

                cel = row.getCell(8);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setSaldo((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setSaldo(valueOf);
                    }
                }

                cel = row.getCell(9);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumeroOrden((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumeroOrden(valueOf);
                    }
                }

                cel = row.getCell(10);
                try {
                    if (cel.getCellType() == CellType.FORMULA) {
                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                            //cell.getNumericCellValue();
                            ingreso.setGuiaChilemat(String.valueOf(cel.getNumericCellValue()).trim());
                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                            //cell.getStringCellValue();
                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                        }
                    } else {
                        try {
                            //cell.getStringCellValue();
//                            System.out.println(cel.getStringCellValue());
                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                        } catch (Exception e) {
//                            System.out.println(cel.getNumericCellValue());
                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                            ingreso.setGuiaChilemat(valueOf);
                        }
                    }
                } catch (Exception ex) {
                    ingreso.setGuiaChilemat("");
                }

                cel = row.getCell(11);
                try {
                    if (cel.getCellType() == CellType.FORMULA) {
                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                            //cell.getNumericCellValue();
                            ingreso.setGuiaProveedor(String.valueOf(cel.getNumericCellValue()).trim());
                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                            //cell.getStringCellValue();
                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                        }
                    } else {
                        try {
                            //cell.getStringCellValue();
//                            System.out.println(cel.getStringCellValue());
                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                        } catch (Exception e) {
//                            System.out.println(cel.getNumericCellValue());
                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                            ingreso.setGuiaProveedor(valueOf);
                        }
                    }
                } catch (Exception ex) {
                    ingreso.setGuiaProveedor("");
                }

                cel = row.getCell(12);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumeroCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumeroCuota(valueOf);
                    }
                }
                arrIngreso.add(ingreso);

                g++;
                row = sheet.getRow(g);
//                System.out.println(g);

                cel = row.getCell(0);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        valueOfx = String.valueOf(cel.getNumericCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        valueOfx = cel.getStringCellValue();
//                        System.out.println("valueOfx " + valueOfx);
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
                        valueOfx = String.valueOf(cel.getStringCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    } catch (Exception e) {
                        //cell.getStringCellValue();
                        valueOfx = String.valueOf(cel.getNumericCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    }
                }
            }
        } else {
            String valueOfx = "ABC";
            while (!"".equals(valueOfx)) {
                DocumentoCobranza ingreso = new DocumentoCobranza();
                cel = row.getCell(0);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumero((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumero(valueOf);
                    }
                }

                cel = row.getCell(1);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setTipo(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setTipo(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setTipo(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setTipo(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(2);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setCuota(valueOf);
                    }
                }

                cel = row.getCell(3);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setSucursal(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setSucursal(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setSucursal(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setSucursal(String.valueOf(valueOf).trim());
                    }
                }
                cel = row.getCell(4);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setProveedor(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setProveedor(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setProveedor(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setProveedor(String.valueOf(valueOf).trim());
                    }
                }
                cel = row.getCell(5);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setFechaEmision(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setFechaEmision(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setFechaEmision(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setFechaEmision(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(6);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setFechaVencimiento(String.valueOf(cel.getNumericCellValue()));
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setFechaVencimiento(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setFechaVencimiento(cel.getStringCellValue());
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setFechaVencimiento(String.valueOf(valueOf).trim());
                    }
                }

                cel = row.getCell(7);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setMontoCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setMontoCuota(valueOf);
                    }
                }

                cel = row.getCell(8);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setSaldo((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setSaldo(valueOf);
                    }
                }

                cel = row.getCell(10);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumeroOrden((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumeroOrden(valueOf);
                    }
                }
                cel = row.getCell(11);
                try {
                    if (cel.getCellType() == CellType.FORMULA) {
                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                            //cell.getNumericCellValue();
                            ingreso.setGuiaChilemat(String.valueOf(cel.getNumericCellValue()).trim());
                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                            //cell.getStringCellValue();
                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                        }
                    } else {
                        try {
                            //cell.getStringCellValue();
//                            System.out.println(cel.getStringCellValue());
                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                        } catch (Exception e) {
//                            System.out.println(cel.getNumericCellValue());
                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                            ingreso.setGuiaChilemat(valueOf);
                        }
                    }
                } catch (Exception ex) {
                    ingreso.setGuiaChilemat("");
                }
                cel = row.getCell(12);
                try {
                    if (cel.getCellType() == CellType.FORMULA) {
                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                            //cell.getNumericCellValue();
                            ingreso.setGuiaProveedor(String.valueOf(cel.getNumericCellValue()).trim());
                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                            //cell.getStringCellValue();
                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                        }
                    } else {
                        try {
                            //cell.getStringCellValue();
//                            System.out.println(cel.getStringCellValue());
                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                        } catch (Exception e) {
//                            System.out.println(cel.getNumericCellValue());
                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                            ingreso.setGuiaProveedor(valueOf);
                        }
                    }
                } catch (Exception ex) {
                    ingreso.setGuiaProveedor("");
                }
                cel = row.getCell(13);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        ingreso.setNumeroCuota((int) cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                        ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                    } catch (Exception e) {
//                        System.out.println(cel.getNumericCellValue());
                        Integer valueOf = (int) cel.getNumericCellValue();
                        ingreso.setNumeroCuota(valueOf);
                    }
                }
                arrIngreso.add(ingreso);

                g++;
                row = sheet.getRow(g);
//                System.out.println(g);

                cel = row.getCell(0);
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        valueOfx = String.valueOf(cel.getNumericCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        valueOfx = cel.getStringCellValue();
//                        System.out.println("valueOfx " + valueOfx);
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
                        valueOfx = String.valueOf(cel.getStringCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    } catch (Exception e) {
                        //cell.getStringCellValue();
                        valueOfx = String.valueOf(cel.getNumericCellValue());
//                        System.out.println("valueOfx " + valueOfx);
                    }
                }
            }
        }

        return arrIngreso;
    }

    public static ArrayList<DocumentoCobranza> leerExcel2(File file) throws FileNotFoundException, IOException {
        ArrayList<DocumentoCobranza> arrIngreso = new ArrayList<>();
        try {
            /*
            int numero;
            String tipo;
            int cuota;
            String proveedor;
            String fechaEmision;
            String fechaVencimiento;
            int montoCuota;
            int saldo;
            int dias;
            int numeroOrden;
            int guiaChilemat;
            int guiaProveedor;
            int numeroCuota;
            String pkNumeroCuota;
             */
            FileInputStream ExcelFileToRead = new FileInputStream(file);
            XSSFWorkbook wb1 = new XSSFWorkbook(ExcelFileToRead);
            XSSFSheet sheet = wb1.getSheetAt(0);

            Row row2;
            Cell cel;
            XSSFRow row = sheet.getRow(12);
            Iterator cells1 = row.cellIterator();
            int cont = 0;
            while (cells1.hasNext()) {
                cel = (Cell) cells1.next();
                if (cel.getCellType() == CellType.FORMULA) {
                    if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                        //cell.getNumericCellValue();
                        cel.getNumericCellValue();
//                        System.out.println(cel.getNumericCellValue());
                    } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                        //cell.getStringCellValue();
                        cel.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                    }
                } else {
                    try {
                        //cell.getStringCellValue();
                        cel.getStringCellValue();
//                        System.out.println(cel.getStringCellValue());
                    } catch (Exception e) {
                        //cell.getNumericCellValue();
                        cel.getNumericCellValue();
//                        System.out.println(cel.getNumericCellValue());
                    }
                }
                cont++;
            }
//            System.out.println("cont " + cont);

            if (cont == 14) {
                Iterator rows1 = sheet.rowIterator();
                int i = 0;
                while (rows1.hasNext()) {
////            System.out.println("i " + i);
                    if (i >= 10) {
                        DocumentoCobranza ingreso = new DocumentoCobranza();
                        row2 = (Row) rows1.next();
                        Iterator cells = row2.cellIterator();
                        cont = 0;
                        while (cells.hasNext()) {
//                            System.out.println(cont);
                            cel = (Cell) cells.next();
                            switch (cont) {
                                case 0:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumero((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumero(valueOf);
                                        }
                                    }
                                    break;
                                case 1:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setTipo(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setTipo(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setTipo(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setTipo(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 2:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 3:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setSucursal(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setSucursal(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setSucursal(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setSucursal(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 4:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setProveedor(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setProveedor(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setProveedor(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setProveedor(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 5:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setFechaEmision(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setFechaEmision(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setFechaEmision(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setFechaEmision(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 6:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setFechaVencimiento(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setFechaVencimiento(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setFechaVencimiento(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setFechaVencimiento(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 7:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setMontoCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setMontoCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 8:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setSaldo((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setSaldo(valueOf);
                                        }
                                    }
                                    break;
                                case 9:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumeroOrden((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumeroOrden(valueOf);
                                        }
                                    }
                                    break;
                                case 10:
                            try {
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getNumericCellValue()).trim());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()).trim());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                                            ingreso.setGuiaChilemat(valueOf);
                                        }
                                    }
                                } catch (Exception ex) {
                                    ingreso.setGuiaChilemat("");
                                }
                                break;
                                case 11:
                            try {
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getNumericCellValue()).trim());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()).trim());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            String valueOf = String.valueOf(cel.getNumericCellValue()).trim();
                                            ingreso.setGuiaProveedor(valueOf);
                                        }
                                    }
                                } catch (Exception ex) {
                                    ingreso.setGuiaProveedor("");
                                }
                                break;
                                case 12:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumeroCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumeroCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 13:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setPkNumeroCuota(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setPkNumeroCuota(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setPkNumeroCuota(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setPkNumeroCuota(String.valueOf(valueOf));
                                        }
                                    }
                                    break;
                            }
                            cont = cont + 1;
                        }
                        arrIngreso.add(ingreso);
                    } else {
                        row2 = (Row) rows1.next();
                    }
                    i++;
//                    System.out.println("i-----------> " + i);
                }
            } else {
                Iterator rows = sheet.rowIterator();
                int i = 0;
                while (rows.hasNext()) {
////            System.out.println("i " + i);
                    if (i >= 10) {
                        DocumentoCobranza ingreso = new DocumentoCobranza();
                        row2 = (Row) rows.next();
                        Iterator cells = row2.cellIterator();
                        cont = 0;
                        while (cells.hasNext()) {
//                            System.out.println(cont);
                            cel = (Cell) cells.next();
                            switch (cont) {
                                case 0:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumero((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumero(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumero(valueOf);
                                        }
                                    }
                                    break;
                                case 1:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setTipo(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setTipo(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setTipo(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setTipo(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 2:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 3:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setSucursal(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setSucursal(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setSucursal(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setSucursal(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 4:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setProveedor(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setProveedor(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setProveedor(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setProveedor(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 5:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setFechaEmision(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setFechaEmision(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setFechaEmision(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setFechaEmision(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 6:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setFechaVencimiento(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setFechaVencimiento(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setFechaVencimiento(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setFechaVencimiento(String.valueOf(valueOf).trim());
                                        }
                                    }
                                    break;
                                case 7:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setMontoCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setMontoCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setMontoCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 8:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setSaldo((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setSaldo(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setSaldo(valueOf);
                                        }
                                    }
                                    break;
                                case 9:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumeroOrden((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumeroOrden(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumeroOrden(valueOf);
                                        }
                                    }
                                    break;
                                case 10:
                            try {
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setGuiaChilemat(String.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            String valueOf = String.valueOf(cel.getNumericCellValue());
                                            ingreso.setGuiaChilemat(valueOf);
                                        }
                                    }
                                } catch (Exception ex) {
                                    ingreso.setGuiaChilemat("");
                                }
                                break;
                                case 11:
                            try {
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setGuiaProveedor(String.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            String valueOf = String.valueOf(cel.getNumericCellValue());
                                            ingreso.setGuiaProveedor(valueOf);
                                        }
                                    }
                                } catch (Exception ex) {
                                    ingreso.setGuiaProveedor("");
                                }
                                break;
                                case 12:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setNumeroCuota((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setNumeroCuota(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setNumeroCuota(valueOf);
                                        }
                                    }
                                    break;
                                case 13:
                                    if (cel.getCellType() == CellType.FORMULA) {
                                        if (cel.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            //cell.getNumericCellValue();
                                            ingreso.setPkNumeroCuota(String.valueOf(cel.getNumericCellValue()));
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setPkNumeroCuota(cel.getStringCellValue());
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setPkNumeroCuota(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setPkNumeroCuota(String.valueOf(valueOf));
                                        }
                                    }
                                    break;
                            }
                            cont = cont + 1;
                        }
                        arrIngreso.add(ingreso);
                    } else {
                        row2 = (Row) rows.next();
                    }
                    i++;
//                    System.out.println("i-----------> " + i);
                }
            }
            return arrIngreso;
        } catch (Exception ex) {
            return arrIngreso;
        }
    }
}
