using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePTGPS.Model
{
     public  class ReportPT
    {
        public int IdPagatodo { get; set; }
        public string Fecha { get; set; }
        public string Mes_Anio { get; set; }
        public string Primeratransaccion { get; set; }
        public string Hora { get; set; }
        public string Tipo { get; set; }
        public int IdSocio { get; set; }
        public string Socio { get; set; }
        public int IdCliente { get; set; }
        public string Cliente { get; set; }
        public string CorreoUsuario { get; set; }
        public string FechaRegistro { get; set; }
        public int Edad { get; set; }
        public string Pais { get; set; }
        public decimal Monto { get; set; }
        public decimal ComisionPagatodo { get; set; }
        public decimal ComisionPagaT { get; set; }
        public decimal ComisionPalpal { get; set; }
        public decimal Total { get; set; }
        public string Resultado { get; set; }
        public string idpaypal { get; set; }
        public string IpRegistro { get; set; }
        public string IpTransaccion { get; set; }
        public int TotalTransacciones { get; set; }
        public int Premovimientos { get; set; }
        public int TotalIPs { get; set; }
        public int totalcard { get; set; }
        public string titular { get; set; }
        public string envio_sms_recarga { get; set; }
        public string certificado { get; set; }
        public string destinatario { get; set; }
        public string Comentarios { get; set; }
        public int Modificaciones { get; set; }
        public int Fraude { get; set; }
        public string Comentarios_internos { get; set; }
        public string Metodo_Pago { get; set; }
        public decimal ComisionMasIVA { get; set; }
    }
}
