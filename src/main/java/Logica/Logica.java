package Logica;

import Clases.DocumentoCobranza;
import DAO.DocumentoCobranzaDAO;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Logica {

    public static void main() throws IOException, SQLException {
        File file = new File("C:\\Users\\sopor\\Desktop\\ChilematNew.xlsx");
//        File file = new File("C:\\Users\\sopor\\Desktop\\COMERCIAL FRANCISCO TOSO LTDA. 06 04 21.xlsx");
        ArrayList<DocumentoCobranza> documentoCobranza = leerExcel(file);
//        System.out.println("");
//        System.out.println(documentoCobranza.size());

        ArrayList<Integer> index = new ArrayList<>();
        AtomicInteger atomicInteger = new AtomicInteger(0);
        documentoCobranza.forEach((DocumentoCobranza doc) -> {
            int numero = doc.getNumero();
            String tipo = doc.getTipo();
            int cuota = doc.getCuota();
            String sucursal = doc.getSucursal();
            String proveedor = doc.getProveedor();
            String fechaEmision = doc.getFechaEmision();
            String fechaVencimiento = doc.getFechaVencimiento();
            int montoCuota = doc.getMontoCuota();
            int saldo = doc.getSaldo();
            int dias = doc.getDias();
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
            } else {
//                System.out.println("------------------  " + valueOf + "   ------------------");
//                System.out.println("numero " + numero);
//                System.out.println("tipo " + tipo);
//                System.out.println("cuota " + cuota);
//                System.out.println("sucursal " + sucursal);
//                System.out.println("proveedor " + proveedor);
//                System.out.println("fechaEmision " + fechaEmision);
//                System.out.println("fechaVencimiento " + fechaVencimiento);
//                System.out.println("montoCuota " + montoCuota);
//                System.out.println("saldo " + saldo);
//                System.out.println("dias " + dias);
//                System.out.println("numeroOrden " + numeroOrden);
//                System.out.println("guiaChilemat " + guiaChilemat);
//                System.out.println("guiaProveedor " + guiaProveedor);
//                System.out.println("numeroCuota " + numeroCuota);
//                System.out.println("pkNumeroCuota " + pkNumeroCuota);
//                System.out.println("----------------------------------------------");
            }
        });

        index.forEach((Integer i) -> {
            documentoCobranza.remove((int) i);
        });

        documentoCobranza.forEach((DocumentoCobranza doc) -> {
            try {
                DocumentoCobranzaDAO.registraDocumentoCobranza(doc, "ingresos");
            } catch (IOException ex) {
                Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
            } catch (SQLException ex) {
                Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

//        System.out.println("leerExcel.size() " + documentoCobranza.size());
    }

    public static ArrayList<DocumentoCobranza> leerExcel(File file) throws FileNotFoundException, IOException {
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
            XSSFSheet sheet1 = wb1.getSheetAt(0);

            Row row2;
            Cell cel;
            XSSFRow row = sheet1.getRow(12);
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
                Iterator rows1 = sheet1.rowIterator();
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
                                            ingreso.setDias((int) cel.getNumericCellValue());
                                        } else if (cel.getCachedFormulaResultType() == CellType.STRING) {
                                            //cell.getStringCellValue();
                                            ingreso.setDias(Integer.valueOf(cel.getStringCellValue()));
                                        }
                                    } else {
                                        try {
                                            //cell.getStringCellValue();
//                                            System.out.println(cel.getStringCellValue());
                                            ingreso.setDias(Integer.valueOf(cel.getStringCellValue()));
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            Integer valueOf = (int) cel.getNumericCellValue();
                                            ingreso.setDias(valueOf);
                                        }
                                    }
                                    break;
                                case 10:
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
                                case 11:
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
                                            ingreso.setGuiaChilemat(cel.getStringCellValue());
                                        } catch (Exception e) {
//                                            System.out.println(cel.getNumericCellValue());
                                            ingreso.setGuiaChilemat(cel.getStringCellValue());
                                        }
                                    }
                                } catch (Exception ex) {
                                    ingreso.setGuiaChilemat("");
                                }
                                break;
                                case 12:
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
                                case 13:
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
                                case 14:
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
                Iterator rows1 = sheet1.rowIterator();

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
                        row2 = (Row) rows1.next();
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
