package DAO;

import Clases.Ingreso;
import Connect.DbConnection;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

public class IngresoDAO {

    public static ArrayList<Ingreso> consultaIngresoVA(String bd) throws IOException, SQLException {
        ArrayList<Ingreso> arrIngreso = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);
        try (PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.ingresova");
                ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                Ingreso ingreso = new Ingreso();
                ingreso.setFolioIngreso(res.getString("folioIngreso"));
                ingreso.setOrdenDeCompra(res.getString("ordenDeCompra"));
                ingreso.setNumeroGuia(res.getString("numeroGuia"));
                ingreso.setAfecto(res.getString("afecto"));
                ingreso.setIva(res.getString("iva"));
                ingreso.setTotal(res.getString("total"));
                ingreso.setProveedor(res.getString("proveedor"));
                ingreso.setAlmacenamiento(res.getString("almacenamineto"));
                ingreso.setEstadoFolio(res.getString("estadoFolio"));
                ingreso.setFechaRecepcion(res.getString("fechaRecepcion"));
                ingreso.setUsuarioIngreso(res.getString("usuarioIngreso"));
                arrIngreso.add(ingreso);
            }
        }
        return arrIngreso;
    }

    public static ArrayList<Ingreso> consultaIngresoOL(String bd) throws IOException, SQLException {
        ArrayList<Ingreso> arrIngreso = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);
        try (PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.ingresool");
                ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                Ingreso ingreso = new Ingreso();
                ingreso.setFolioIngreso(res.getString("folioIngreso"));
                ingreso.setOrdenDeCompra(res.getString("ordenDeCompra"));
                ingreso.setNumeroGuia(res.getString("numeroGuia"));
                ingreso.setAfecto(res.getString("afecto"));
                ingreso.setIva(res.getString("iva"));
                ingreso.setTotal(res.getString("total"));
                ingreso.setProveedor(res.getString("proveedor"));
                ingreso.setAlmacenamiento(res.getString("almacenamineto"));
                ingreso.setEstadoFolio(res.getString("estadoFolio"));
                ingreso.setFechaRecepcion(res.getString("fechaRecepcion"));
                ingreso.setUsuarioIngreso(res.getString("usuarioIngreso"));
                arrIngreso.add(ingreso);
            }
        }
        return arrIngreso;
    }

    public static ArrayList<Ingreso> consultaIngresoPB(String bd) throws IOException, SQLException {
        ArrayList<Ingreso> arrIngreso = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);
        try (PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.ingresopb");
                ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                Ingreso ingreso = new Ingreso();
                ingreso.setFolioIngreso(res.getString("folioIngreso"));
                ingreso.setOrdenDeCompra(res.getString("ordenDeCompra"));
                ingreso.setNumeroGuia(res.getString("numeroGuia"));
                ingreso.setAfecto(res.getString("afecto"));
                ingreso.setIva(res.getString("iva"));
                ingreso.setTotal(res.getString("total"));
                ingreso.setProveedor(res.getString("proveedor"));
                ingreso.setAlmacenamiento(res.getString("almacenamineto"));
                ingreso.setEstadoFolio(res.getString("estadoFolio"));
                ingreso.setFechaRecepcion(res.getString("fechaRecepcion"));
                ingreso.setUsuarioIngreso(res.getString("usuarioIngreso"));
                arrIngreso.add(ingreso);
            }
        }
        return arrIngreso;
    }
}
