package DAO;

import Connect.DbConnection;
import Clases.DocumentoCobranza;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import javax.swing.JOptionPane;

public class DocumentoCobranzaDAO {

    public static void registraDocumentoCobranza(DocumentoCobranza documentoCobranza, String bd) throws IOException, SQLException {
        try {
            DbConnection conex = new DbConnection(bd);
            Statement estatuto = conex.getConnection().createStatement();
            estatuto.executeUpdate("INSERT INTO ingresos.cobranzaChilemat (numero, tipo, cuota, sucursal, proveedor, fechaEmision, fechaVencimiento, montoCuota, saldo, dias, numeroOrden, guiaChilemat, guiaProveedor, numeroCuota, comentarioNotaDeCredito, pkNumeroCuota) VALUES ('"
                    + documentoCobranza.getNumero() + "', '"
                    + documentoCobranza.getTipo() + "', '"
                    + documentoCobranza.getCuota() + "', '"
                    + documentoCobranza.getSucursal() + "', '"
                    + documentoCobranza.getProveedor() + "', '"
                    + documentoCobranza.getFechaEmision() + "', '"
                    + documentoCobranza.getFechaVencimiento() + "', '"
                    + documentoCobranza.getMontoCuota() + "', '"
                    + documentoCobranza.getSaldo() + "', '" + 0 + "', '"
                    + documentoCobranza.getNumeroOrden() + "', '"
                    + documentoCobranza.getGuiaChilemat().trim() + "', '"
                    + documentoCobranza.getGuiaProveedor().trim() + "', '"
                    + documentoCobranza.getNumeroCuota() + "', '"
                    + documentoCobranza.getComentarioNotaDeCrefito() + "', '"
                    + documentoCobranza.getPkNumeroCuota() + "')"
            );
            estatuto.close();
            conex.desconectar();
        } catch (Exception ex) {

        }
    }

    public static void actualizaDocumentoCobranzaGuiaChilemat(DocumentoCobranza documentoCobranza, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.cobranzachilemat SET"
                    + " numero = '" + documentoCobranza.getNumero() + "' ,"
                    + " tipo = '" + documentoCobranza.getTipo() + "' ,"
                    + " cuota = '" + documentoCobranza.getCuota() + "' ,"
                    + " sucursal = '" + documentoCobranza.getSucursal() + "' ,"
                    + " proveedor = '" + documentoCobranza.getProveedor() + "' ,"
                    + " fechaEmision = '" + documentoCobranza.getFechaEmision() + "' ,"
                    + " fechaVencimiento = '" + documentoCobranza.getFechaVencimiento() + "' ,"
                    + " montoCuota = '" + documentoCobranza.getMontoCuota() + "' ,"
                    + " saldo = '" + documentoCobranza.getSaldo() + "' ,"
                    + " dias = '" + documentoCobranza.getDias() + "' ,"
                    + " numeroOrden = '" + documentoCobranza.getNumeroOrden() + "' ,"
                    + " guiaChilemat = '" + documentoCobranza.getGuiaChilemat() + "' ,"
                    + " numeroCuota = '" + documentoCobranza.getNumeroCuota() + "'"
                    + " WHERE pkNumeroCuota = '" + documentoCobranza.getPkNumeroCuota() + "'");
            conex.desconectar();
        }
    }

    public static void actualizaDocumentoCobranzaGuiaProveedor(DocumentoCobranza documentoCobranza, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);

        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.cobranzachilemat SET"
                    + " numero = '" + documentoCobranza.getNumero() + "' ,"
                    + " tipo = '" + documentoCobranza.getTipo() + "' ,"
                    + " cuota = '" + documentoCobranza.getCuota() + "' ,"
                    + " sucursal = '" + documentoCobranza.getSucursal() + "' ,"
                    + " proveedor = '" + documentoCobranza.getProveedor() + "' ,"
                    + " fechaEmision = '" + documentoCobranza.getFechaEmision() + "' ,"
                    + " fechaVencimiento = '" + documentoCobranza.getFechaVencimiento() + "' ,"
                    + " montoCuota = '" + documentoCobranza.getMontoCuota() + "' ,"
                    + " saldo = '" + documentoCobranza.getSaldo() + "' ,"
                    + " dias = '" + documentoCobranza.getDias() + "' ,"
                    + " numeroOrden = '" + documentoCobranza.getNumeroOrden() + "' ,"
                    + " guiaProveedor = '" + documentoCobranza.getGuiaProveedor() + "' ,"
                    + " numeroCuota = '" + documentoCobranza.getNumeroCuota() + "'"
                    + " WHERE pkNumeroCuota = '" + documentoCobranza.getPkNumeroCuota() + "'");
            conex.desconectar();
        }
    }

    public static ArrayList<DocumentoCobranza> consultaDocumentoCobranza(String bd) throws IOException, SQLException {
        ArrayList<DocumentoCobranza> arrDocumentoCobranza = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);

        try ( PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.cobranzaChilemat");  ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                DocumentoCobranza doc = new DocumentoCobranza();
                doc.setNumero(res.getInt("numero"));
                doc.setTipo(res.getString("tipo"));
                doc.setCuota(res.getInt("cuota"));
                doc.setSucursal(res.getString("sucursal"));
                doc.setProveedor(res.getString("proveedor"));
                doc.setFechaEmision(res.getString("fechaEmision"));
                doc.setFechaVencimiento(res.getString("fechaVencimiento"));
                doc.setMontoCuota(res.getInt("montoCuota"));
                doc.setSaldo(res.getInt("saldo"));
                doc.setDias(0);
                doc.setNumeroOrden(res.getInt("numeroOrden"));
                doc.setGuiaChilemat(res.getString("guiaChilemat"));
                doc.setGuiaProveedor(res.getString("guiaProveedor"));
                doc.setNumeroCuota(res.getInt("numeroCuota"));
                doc.setPkNumeroCuota(res.getString("pkNumeroCuota"));
                doc.setComentario(res.getString("comentario"));
                doc.setEstado(res.getString("estado"));
                doc.setComentarioNotaDeCrefito(res.getInt("comentarioNotaDeCredito"));
                arrDocumentoCobranza.add(doc);
            }
        }
        return arrDocumentoCobranza;
    }

    public static ArrayList<DocumentoCobranza> consultaDocumentoCobranza2(String bd) throws IOException, SQLException {
        ArrayList<DocumentoCobranza> arrDocumentoCobranza = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);

        try ( PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.cobranzaChilemat2");  ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                DocumentoCobranza doc = new DocumentoCobranza();
                doc.setNumero(res.getInt("numero"));
                doc.setTipo(res.getString("tipo"));
                doc.setCuota(res.getInt("cuota"));
                doc.setSucursal(res.getString("sucursal"));
                doc.setProveedor(res.getString("proveedor"));
                doc.setFechaEmision(res.getString("fechaEmision"));
                doc.setFechaVencimiento(res.getString("fechaVencimiento"));
                doc.setMontoCuota(res.getInt("montoCuota"));
                doc.setSaldo(res.getInt("saldo"));
                doc.setDias(0);
                doc.setNumeroOrden(res.getInt("numeroOrden"));
                doc.setGuiaChilemat(res.getString("guiaChilemat"));
                doc.setGuiaProveedor(res.getString("guiaProveedor"));
                doc.setNumeroCuota(res.getInt("numeroCuota"));
                doc.setPkNumeroCuota(res.getString("pkNumeroCuota"));
                doc.setComentario(res.getString("comentario"));
                doc.setEstado(res.getString("estado"));
                doc.setComentarioNotaDeCrefito(res.getInt("comentarioNotaDeCredito"));
                arrDocumentoCobranza.add(doc);
            }
        }
        return arrDocumentoCobranza;
    }

    public static DocumentoCobranza consultaDocumentoCobranzaUnico(String pkNumeroCuota, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        DocumentoCobranza doc = new DocumentoCobranza();

        try ( PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.cobranzaChilemat WHERE pkNumeroCuota = '" + pkNumeroCuota + "'");  ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                doc.setNumero(res.getInt("numero"));
                doc.setTipo(res.getString("tipo"));
                doc.setCuota(res.getInt("cuota"));
                doc.setSucursal(res.getString("sucursal"));
                doc.setProveedor(res.getString("proveedor"));
                doc.setFechaEmision(res.getString("fechaEmision"));
                doc.setFechaVencimiento(res.getString("fechaVencimiento"));
                doc.setMontoCuota(res.getInt("montoCuota"));
                doc.setSaldo(res.getInt("saldo"));
                doc.setDias(0);
                doc.setNumeroOrden(res.getInt("numeroOrden"));
                doc.setGuiaChilemat(res.getString("guiaChilemat"));
                doc.setGuiaProveedor(res.getString("guiaProveedor"));
                doc.setNumeroCuota(res.getInt("numeroCuota"));
                doc.setPkNumeroCuota(res.getString("pkNumeroCuota"));
                doc.setComentario(res.getString("comentario"));
                doc.setEstado(res.getString("estado"));
                doc.setComentarioNotaDeCrefito(res.getInt("comentarioNotaDeCredito"));
            }
        }
        return doc;
    }

    public static void actualizaComentarioDocumentoCobranza(String comentario, String pknumerocuota, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.cobranzachilemat SET "
                    + "comentario = '" + comentario
                    + "' WHERE pkNumeroCuota = '" + pknumerocuota + "'");
            conex.desconectar();
        }
    }

    public static boolean actualizaComentarioNotaDeCredito(String comentario, String pknumerocuota, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.cobranzachilemat SET "
                    + "comentarioNotaDeCredito = '" + comentario
                    + "' WHERE pkNumeroCuota = '" + pknumerocuota + "'");
            conex.desconectar();
            return true;
        } catch (Exception ex) {
            return false;
        }
    }

    public static void actualizaEstadoDocumentoCobranza(String estado, String pknumerocuota, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.cobranzachilemat SET "
                    + "estado = '" + estado
                    + "' WHERE pkNumeroCuota = '" + pknumerocuota + "'");
            conex.desconectar();
        }
    }
}
