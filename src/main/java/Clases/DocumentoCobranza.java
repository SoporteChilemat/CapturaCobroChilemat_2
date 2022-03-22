package Clases;

public class DocumentoCobranza {

    int numero;
    String tipo;
    int cuota;
    String sucursal;
    String proveedor;
    String fechaEmision;
    String fechaVencimiento;
    int montoCuota;
    int saldo;
    int dias;
    int numeroOrden;
    String guiaChilemat;
    String guiaProveedor;
    int numeroCuota;
    String pkNumeroCuota;
    String comentario;
    String estado;
    int comentarioNotaDeCrefito;

    public int getComentarioNotaDeCrefito() {
        return comentarioNotaDeCrefito;
    }

    public void setComentarioNotaDeCrefito(int comentarioNotaDeCrefito) {
        this.comentarioNotaDeCrefito = comentarioNotaDeCrefito;
    }

    public String getEstado() {
        return estado;
    }

    public void setEstado(String estado) {
        this.estado = estado;
    }

    public String getComentario() {
        return comentario;
    }

    public void setComentario(String comentario) {
        this.comentario = comentario;
    }

    public int getNumero() {
        return numero;
    }

    public void setNumero(int numero) {
        this.numero = numero;
    }

    public String getTipo() {
        return tipo;
    }

    public void setTipo(String tipo) {
        this.tipo = tipo;
    }

    public int getCuota() {
        return cuota;
    }

    public void setCuota(int cuota) {
        this.cuota = cuota;
    }

    public String getSucursal() {
        return sucursal;
    }

    public void setSucursal(String sucursal) {
        this.sucursal = sucursal;
    }

    public String getProveedor() {
        return proveedor;
    }

    public void setProveedor(String proveedor) {
        this.proveedor = proveedor;
    }

    public String getFechaEmision() {
        return fechaEmision;
    }

    public void setFechaEmision(String fechaEmision) {
        this.fechaEmision = fechaEmision;
    }

    public String getFechaVencimiento() {
        return fechaVencimiento;
    }

    public void setFechaVencimiento(String fechaVencimiento) {
        this.fechaVencimiento = fechaVencimiento;
    }

    public int getMontoCuota() {
        return montoCuota;
    }

    public void setMontoCuota(int montoCuota) {
        this.montoCuota = montoCuota;
    }

    public int getSaldo() {
        return saldo;
    }

    public void setSaldo(int saldo) {
        this.saldo = saldo;
    }

    public int getDias() {
        return dias;
    }

    public void setDias(int dias) {
        this.dias = dias;
    }

    public int getNumeroOrden() {
        return numeroOrden;
    }

    public void setNumeroOrden(int numeroOrden) {
        this.numeroOrden = numeroOrden;
    }

    public int getNumeroCuota() {
        return numeroCuota;
    }

    public void setNumeroCuota(int numeroCuota) {
        this.numeroCuota = numeroCuota;
    }

    public String getPkNumeroCuota() {
        return pkNumeroCuota;
    }

    public void setPkNumeroCuota(String pkNumeroCuota) {
        this.pkNumeroCuota = pkNumeroCuota;
    }

    public String getGuiaChilemat() {
        return guiaChilemat;
    }

    public void setGuiaChilemat(String guiaChilemat) {
        this.guiaChilemat = guiaChilemat;
    }

    public String getGuiaProveedor() {
        return guiaProveedor;
    }

    public void setGuiaProveedor(String guiaProveedor) {
        this.guiaProveedor = guiaProveedor;
    }
}
