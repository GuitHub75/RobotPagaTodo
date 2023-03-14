using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ReportePT;
using ReportePTGPS.Model;
using Dapper;
using ClosedXML.Excel;

namespace ReportePTGPS
{
     public class Metodos
    {
        Connection _con = new Connection();
        ConnectionPT _conPT = new ConnectionPT();
        public MemoryStream GetStream(XLWorkbook excelWorkbook)
        {
            MemoryStream fs = new MemoryStream();
            excelWorkbook.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }

        #region "FUNCIONES PRINCIPALES"
        public int SendDataEmailWeekEnd()
        {
            MemoryStream[] files1 = new MemoryStream[3];
            MemoryStream[] files2 = new MemoryStream[3];
            files1 = ReportePTRocketWeekEnd();
            files2 = ReportePTtWeekEnd();

            Correo _correo = new Correo();
            _correo.SMTP4_1(files1[0], files1[1], files1[2], files2[0], files2[1], files2[2]);
            return 1;
        }

        public int SendDataEmaiYesterday()
        {
            MemoryStream files1 = new MemoryStream();
            MemoryStream files2 = new MemoryStream();
            files1 = ReportePTRocket();
            files2 = ReportePT();

            Correo _correo = new Correo();
            _correo.SMTP4_2(files1, files2);
            return 1;
        }
        #endregion

        #region "FUNCIONES RETORNAN NUESTROS ARCHIVOS CON SU DATA"

        //ESTE ES EL REPORTE PT DE ROCKET DE DIA VIERNES, SÁBADO Y DOMINGO 
        public MemoryStream[] ReportePTRocketWeekEnd()
        {
            MemoryStream[] files = new MemoryStream[3];
            List<ReportPTRocket> _PTRocketFriday = new List<ReportPTRocket>();
            List<ReportPTRocket> _PTRocketSaturday = new List<ReportPTRocket>();
            List<ReportPTRocket> _PTRocketSunday = new List<ReportPTRocket>();
            try
            {

                _PTRocketFriday = GetReportList(3, 2);
                _PTRocketSaturday = GetReportList(2, 1);
                _PTRocketSunday = GetReportList(1, 0);

                DataTable dtFriday = new DataTable();
                DataTable dtSaturday = new DataTable();
                DataTable dtSunday = new DataTable();

                dtFriday = setDataTable();
                dtSaturday = setDataTable();
                dtSunday = setDataTable();

                MemoryStream stream = GetMemoryStream(_PTRocketFriday, dtFriday);
                MemoryStream stream2 = GetMemoryStream(_PTRocketSaturday, dtSaturday);
                MemoryStream stream3 = GetMemoryStream(_PTRocketSunday, dtSunday);

                files[0] = stream;
                files[1] = stream2;
                files[2] = stream3;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return files;
        }
        //ESTE ES EL REPORTE PT DE DIA VIERNES, SÁBADO Y DOMINGO 
        public MemoryStream[] ReportePTtWeekEnd()
        {
            MemoryStream[] files = new MemoryStream[3];
            List<ReportPT> _PTFriday = new List<ReportPT>();
            List<ReportPT> _PTSaturday = new List<ReportPT>();
            List<ReportPT> _PTSunday = new List<ReportPT>();
            try
            {
                _PTFriday = GetReportListPT(3, 2);
                _PTSaturday = GetReportListPT(2, 1);
                _PTSunday = GetReportListPT(1, 0);

                DataTable dtFriday = new DataTable();
                DataTable dtSaturday = new DataTable();
                DataTable dtSunday = new DataTable();

                dtFriday = setDataTablePT();
                dtSaturday = setDataTablePT();
                dtSunday = setDataTablePT();


                MemoryStream stream = GetMemoryStreamPT(_PTFriday, dtFriday);
                MemoryStream stream2 = GetMemoryStreamPT(_PTSaturday, dtSaturday);
                MemoryStream stream3 = GetMemoryStreamPT(_PTSunday, dtSunday);

                files[0] = stream;
                files[1] = stream2;
                files[2] = stream3;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return files;
        }

        //ESTE ES EL REPORTE PT ROCKET  DIARIO 
        public MemoryStream ReportePTRocket()
        {
            MemoryStream file = new MemoryStream();
            try
            {
                List<ReportPTRocket> ReportPTyesterday = new List<ReportPTRocket>();
                ReportPTyesterday = GetReportList(1, 0);

                DataTable dtyesterday = new DataTable();

                dtyesterday = setDataTable();

                file = GetMemoryStream(ReportPTyesterday, dtyesterday);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return file;
        }
        //ESTE ES EL REPORTE PT DIARIO
        public MemoryStream ReportePT()
        {
            MemoryStream file = new MemoryStream();
            try
            {
                List<ReportPT> ReportPTyesterday = new List<ReportPT>();
                ReportPTyesterday = GetReportListPT(1, 0);

                DataTable dtyesterday = new DataTable();

                dtyesterday = setDataTablePT();

                file = GetMemoryStreamPT(ReportPTyesterday, dtyesterday);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return file;
        }
        #endregion

        #region "FUNCIONES RETORNAN LAS LISTAS CON SU RESPECTIVA DATA"
        public List<ReportPTRocket> GetReportList(int day1, int day2)
        {
            List<ReportPTRocket> _ListReport1 = new List<ReportPTRocket>();
            List<ReportPTRocket> _ListReport2 = new List<ReportPTRocket>();
            List<ReportPTRocket> _ListReport3 = new List<ReportPTRocket>();
            try
            {
                _con.Cnn.Open();
                _ListReport1 = _con.Cnn.Query<ReportPTRocket>(@"
                              select
                                       prs.PaymentServiceID as PaymentServiceID,
                                        FORMAT(DATEADD(Hh, -6, prs.[RegDate]),'dd/MM/yyyy')+' '+ FORMAT(DATEADD(Hh, -6, prs.[RegDate])  ,N'hh:mm tt')as RegDate,
                                       spc.Name as ServiceProviderCategory, 
                                       sp.SocioID as SocioID, 
                                       sp.Name as NameServiceProvider, 
                                       prs.PersonID as PersonID,
                                       prs.UserName as ReceiverName,
                                       prs.Amount as Amount,
                                       prs.CommissionWithheld as Comission, 
                                       prs.Amount + prs.CommissionWithheld as TotalWithComission,
                                       prs.Status as Status,
                                       prs.ServiceProviderUserID as Receiver,
                                       tp.Reference as Identifier,
									   cc.CreditCardNumber
                    --prs.ServiceProviderID as ServiceProviderID,
                    --prs.ExpirationDate as ExpirationDate, 
                    --pea.Email as Email, 
                    --pp.FirstName + ' ' + pp.LastName as PersonName, 
                    --prs.RegDate as SerfinsaDate,
                    --prs.Amount as TotalAT,
                    --prs.Amount as TotalCharged, 
                    --DATEADD(hh, -6, prs.[RegDate]) as TransactionDate,														
                    --tp.CreditCardID as CreditCardID, tp.[Authorization] as Numero,
                    
               from Rocket.PaymentService prs
           INNER JOIN Rocket.ServiceProvider sp ON sp.ServiceProviderID = prs.ServiceProviderID
           INNER JOIN Person.Person pp ON pp.PersonID = prs.PersonID
           INNER JOIN Person.EmailAddress pea ON pea.PersonID = prs.PersonID
           INNER JOIN Rocket.ServiceProviderCategory spc ON spc.ServiceProviderCategoryID = sp.ServiceProviderCategoryID
           INNER JOIN Rocket.TransactionPOS tp ON tp.TransactionPOSID = prs.TransactionPOSID
           INNER JOIN Person.CreditCard cc ON cc.CreditCardID = tp.CreditCardID
       where (cc.RocketCard = 0 OR cc.RocketCard IS NULL) 
         AND DATEADD(hh, -6, prs.[RegDate]) 
               between
			    FORMAT(dateadd(hh,-6, dateadd(dd,-@day1, getdate())),'yyyy-MM-dd 00:00:00')  
			   and
			    FORMAT(dateadd(hh,-6, dateadd(dd,-@day2, getdate())),'yyyy-MM-dd 00:00:00')
         AND prs.Status = 1
         AND prs.PaymentConfirmation = 1
         AND spc.ServiceProviderCategoryID = 2
                 Order By PaymentServiceID DESC

                       ", new { day1 = day1, day2 = day2 }).ToList();
                _ListReport2 = _con.Cnn.Query<ReportPTRocket>(@"
                              select 
                         prs.PaymentServiceID as PaymentServiceID, 
						 FORMAT(DATEADD(Hh, -6, prs.[RegDate]),'dd/MM/yyyy')+' '+ FORMAT(DATEADD(Hh, -6, prs.[RegDate])  ,N'hh:mm tt')as RegDate,
						 spc.Name as ServiceProviderCategory, 
						 sp.SocioID as SocioID, 
						 sp.Name as NameServiceProvider,
						 prs.PersonID as PersonID,
						 prs.UserName as ReceiverName,
						 prs.Amount as Amount,
						 prs.CommissionWithheld as Comission,
						 prs.Amount + prs.CommissionWithheld as TotalWithComission,
						 prs.Status as Status,
						 prs.ServiceProviderUserID as Receiver,
						 tp.Reference as Identifier,
						 cc.CreditCardNumber

						 --prs.ServiceProviderID as ServiceProviderID,
						 --prs.ExpirationDate as ExpirationDate,
						 -- pp.FirstName + ' ' + pp.LastName as PersonName, 
						 -- pea.Email as Email,
						 -- spc.Name as ServiceProviderCategory, 
                         -- prs.RegDate as SerfinsaDate, 
						 -- prs.Amount as TotalAT,
						 -- prs.Amount as TotalCharged, 
						 -- DATEADD(hh, -6, prs.[RegDate]) as TransactionDate,
						 -- tp.CreditCardID as CreditCardID,
						 -- tp.[Authorization] as Numer
						 
                                                            from Rocket.PaymentService prs
                                                            INNER JOIN Rocket.ServiceProvider sp ON sp.ServiceProviderID = prs.ServiceProviderID
                                                            INNER JOIN Person.Person pp ON pp.PersonID = prs.PersonID
                                                            INNER JOIN Person.EmailAddress pea ON pea.PersonID = prs.PersonID
                                                            INNER JOIN Rocket.ServiceProviderCategory spc ON spc.ServiceProviderCategoryID = sp.ServiceProviderCategoryID
                                                            INNER JOIN Rocket.TransactionPOS tp ON tp.TransactionPOSID = prs.TransactionPOSID
															INNER JOIN Person.CreditCard cc ON cc.CreditCardID = tp.CreditCardID
															INNER JOIN Person.RocketCardVolcan cr on cr.CreditCardID = cc.CreditCardID
                                                            where DATEADD(hh, -6, prs.[RegDate])
															between
			                                                   FORMAT(dateadd(hh,-6, dateadd(dd,-@day1, getdate())),'yyyy-MM-dd 00:00:00')  
			                                                              and
			                                                      FORMAT(dateadd(hh,-6, dateadd(dd,-@day2, getdate())),'yyyy-MM-dd 00:00:00')
															AND prs.Status = 1
                                                            AND prs.PaymentConfirmation = 1
															AND spc.ServiceProviderCategoryID = 2
                                                            Order By PaymentServiceID DESC

                       ", new { day1 = day1, day2 = day2 }).ToList();
                _con.Cnn.Close();

                _ListReport3 = MatchList(_ListReport1, _ListReport2);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return _ListReport3;
        }
        public List<ReportPT> GetReportListPT(int day1, int day2)
        {
            List<ReportPT> _ListReport1 = new List<ReportPT>();
            List<ReportPT> _ListReport2 = new List<ReportPT>();
            List<ReportPT> _ListReport3 = new List<ReportPT>();
            try
            {
                _conPT.Cnn.Open();
                _ListReport1 = _conPT.Cnn.Query<ReportPT>(@"
                  select
                  pt.key_id as IdPagatodo,
                  FORMAT(DATEADD(hh,-6, pt.regdate),'dd/MM/yyyy') as Fecha,
                  FORMAT(DATEADD(hh,-6, pt.regdate),'MMM-yy') as Mes_Anio, 
                  (
				  SELECT  min(DATEADD(Hh,-6,mxc.regdate)) FROM
				  PT_MovimientosXCuenta AS mxc 
                    INNER JOIN PT_CuentasXCliente AS pxc on pxc.key_id = mxc.id_cuenta 
				  WHERE pxc.id_cliente=pt3.key_id
				  )  as Primeratransaccion,
                 FORMAT(DATEADD(hh,-6, pt.regdate),'HH:mm:ss') as Hora, 
                 (case pt.tipo_movimiento when 1 then 'Pago' when 3 then 'Certificado de Regalo' else '----' end) as Tipo,
                  pt2.idSocio as IdSocio,
				  pt4.nombre as Socio,
				  pt3.key_id as IdCliente,
                 (pt3.nombres + ' ' + pt3.apellidos) as Cliente, 
                 pt3.email as CorreoUsuario, 
                 FORMAT(DATEADD(hh, -6, pt3.regdate),'dd/MM/yyyy') as FechaRegistro, 
                 DATEDIFF(YEAR, pt3.fecha_nacimiento,dateadd(hh,-6,GETDATE())) as Edad, 
			     pt5.nombre as Pais, convert(float,pt.monto) as Monto,
				 convert(float,pt.comision) as ComisionPagatodo,
				 convert(float,pt.comision_paypal) as ComisionPagaT,
				 convert(float,pt.comision_paypal_fee) as ComisionPalpal,
				 convert(float,pt.total ) as Total,
                 'EXITOSO' as Resultado, 
				 pt.id_paypal as idpaypal,
				 pt3.first_ip as IpRegistro,
				 pt.dir_ip as IpTransaccion,
                (select COUNT(mxc.key_id) 
				from dbo.PT_MovimientosXCuenta mxc 
                    inner join dbo.PT_CuentasXCliente cxc on cxc.key_id = mxc.id_cuenta
                    where cxc.id_cliente = pt3.key_id and mxc.estado_paypal = 1 and mxc.estado_pago_socio = 1)
					as TotalTransacciones, 
                (select COUNT(pre.key_id) from dbo.PT_PreMovimientosXCuenta pre 
                where pre.idCliente = pt3.key_id)
				as Premovimientos,
				(select count( distinct pre.dir_ip) from dbo.PT_PreMovimientosXCuenta pre 
                where pre.idCliente = pt3.key_id)
				as TotalIPs,
                 (SELECT count(distinct LTRIM(RTRIM(SUBSTRING(mc.descripcion, CHARINDEX('<cc-number>',mc.descripcion,0)+11,16)))) 
                FROM PT_MovimientosXCuenta mc inner join PT_PreMovimientosXCuenta pmc on mc.idPreMovimiento=pmc.key_id 
               WHERE pmc.idCliente= pt3.key_id and tipo_pago_cliente=2)
		  as totalcard,
               (case pt.tipo_movimiento when 2 then '----' else 
		       (SELECT pxc.detalle FROM PT_MovimientosXCuenta mxc inner join PT_PreMovimientosXCuenta pxc on mxc.idPreMovimiento=pxc.key_id
		        WHERE mxc.key_id=pt.key_id) end)
		   as titular,
               (SELECT CASE  pt.[envio_sms_recarga] 
			   WHEN 1 THEN Convert(varchar,(SELECT v.valor FROM [dbo].[PT_CambioValorSMS] v WHERE pt.regdate >= v.regdate and pt.regdate <= v.moddate)) 
	           WHEN 0 THEN 'N/A' ELSE 'N/A' END)
			AS envio_sms_recarga,
                    CASE pt.tipo_movimiento WHEN 1 THEN '---' WHEN 2 THEN '---' ELSE pt6.codigo_certificado END
			AS certificado,
            pt.destinatario,
			pt.Comentarios,
			pt.Modificaciones,
			pt.fraude,
			pt.Comentarios_internos,
               CASE pt.tipo_pago_cliente WHEN 2 THEN case len(pt.id_paypal)
		      when 6 then 'SERFINSA' else 'CREDOMATIC'  end  WHEN 1 THEN 'PAYPAL' END
		  AS Metodo_Pago,
          CONVERT(FLOAT, pt.total * 3/100 )as ComisionMasIVA
 
                      from dbo.PT_MovimientosXCuenta pt
                      inner join dbo.PT_PreMovimientosXCuenta pt2 on pt2.key_id = pt.idPreMovimiento
                       inner join dbo.PT_Clientes pt3 on pt3.key_id = pt2.idCliente
                       inner join dbo.PT_Socios pt4 on pt4.key_id = pt2.idSocio
                       inner join dbo.PT_Pais pt5 on pt5.key_id = pt3.pais
                       left join dbo.PT_Certificados pt6 on pt6.id_movimiento = pt.key_id
                       where pt.estado_paypal = 1 and pt.estado_pago_socio = 1 and (pt.tipo_movimiento = 1 or pt.tipo_movimiento = 3) 
                       and DATEADD(hh,-6, pt.regdate) >= FORMAT(dateadd(hh,-6, dateadd(dd,-@day1, getdate())),'yyyy-MM-dd 00:00:00')
                       and DATEADD(hh,-6, pt.regdate) <= FORMAT(dateadd(hh,-6, dateadd(dd,-@day2, getdate())),'yyyy-MM-dd 00:00:00')
                       order by pt.regdate desc

                       ", new { day1 = day1, day2 = day2 }).ToList();
                _conPT.Cnn.Close();
            }
            catch (Exception EX)
            {
                Console.WriteLine(EX);
            }
            return _ListReport1;
        }
        #endregion

        #region "ESTOS SON METODOS ESPECIALES ROCKET"
        //Con esta funcion hacemos match de dos listas que extraemos de base de datos
        public List<ReportPTRocket> MatchList(List<ReportPTRocket> _list1, List<ReportPTRocket> _list2)
        {
            List<ReportPTRocket> _list = new List<ReportPTRocket>();
            try
            {
                foreach (var item in _list1)
                {
                    item.Identifier = item.Identifier + Model.AESEncrytDecry.DecryptStringAES1(item.CreditCardNumber);
                    _list.Add(item);
                }
                foreach (var item in _list2)
                {
                    item.Identifier = item.Identifier + item.CreditCardNumber;
                    _list.Add(item);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return _list;
        }
        public DataTable setDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[13] {
                    new DataColumn("PaymentServiceID", typeof(string)),
                    new DataColumn("RegDate", typeof(string)),
                    new DataColumn("ServiceProviderCategory", typeof(string)),
                    new DataColumn("SocioID", typeof(string)),
                    new DataColumn("NameServiceProvider", typeof(string)),
                    new DataColumn("PersonID", typeof(string)),
                    new DataColumn("ReceiverName", typeof(string)),
                    new DataColumn("Amount", typeof(decimal)),
                    new DataColumn("Comission", typeof(decimal)),
                    new DataColumn("TotalWithComission", typeof(decimal)),
                    new DataColumn("Status", typeof(string)),
                    new DataColumn("Receiver", typeof(string)),
                    new DataColumn("Identifier", typeof(string))
                });
            return dt;

        }
        public MemoryStream GetMemoryStream(List<ReportPTRocket> lista, DataTable dt)
        {
            decimal TotalAmount = 0;
            decimal TotalComission = 0;
            decimal TotalWidthComission = 0;

            foreach (var item1 in lista)
            {
                TotalAmount += item1.Amount;
                TotalComission += item1.Comission;
                TotalWidthComission += item1.TotalWithComission;
                string Status = "-";
                if (item1.ServiceProviderCategory == "Servicios")
                {
                    item1.ServiceProviderCategory = "Pago";
                }

                if (item1.Status == 1)
                {
                    Status = "Exitoso";
                }
                dt.Rows.Add(
                    item1.PaymentServiceID,
                    item1.RegDate,
                    item1.ServiceProviderCategory,
                    item1.SocioID,
                    item1.NameServiceProvider,
                    item1.PersonID,
                    item1.ReceiverName,
                    item1.Amount,
                    item1.Comission,
                    item1.TotalWithComission,
                    Status,
                    item1.Receiver,
                    item1.Identifier
                    );
            }
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "TOTAL", TotalAmount, TotalComission, TotalWidthComission, "-", "-", "-");



            MemoryStream stream;
            using (XLWorkbook wb = new XLWorkbook())
            {
                var rs = wb.Worksheets.Add("ReportePTRocket");
              

                rs.Range("A2:M2").Style.Font.SetBold();
                rs.Range("A2:M2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rs.Range("A2:M2").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.PowderBlue;

                rs.Cell("A2").Value = "Id Paga Todo";
                rs.Cell("B2").Value = "Fecha";
                rs.Cell("C2").Value = "Tipo";
                rs.Cell("D2").Value = "Id Socio";
                rs.Cell("E2").Value = "Socio";
                rs.Cell("F2").Value = "Id Cliente";
                rs.Cell("G2").Value = "Cliente";
                rs.Cell("H2").Value = "Monto";
                rs.Cell("I2").Value = "Comision PagaTodo";
                rs.Cell("J2").Value = "Total";
                rs.Cell("K2").Value = "Resultado";
                rs.Cell("L2").Value = "Destinatario";
                rs.Cell("M2").Value = "Identificador";

                rs.Cell("A3").InsertData(dt.Rows);

                rs.CellsUsed().Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.BottomBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.TopBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.LeftBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.RightBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rs.Columns().AdjustToContents();
               
                stream = GetStream(wb); //Método se define arriba

            }

            return stream;
        }

        #endregion

        #region "ESTOS SON METODOS ESPECIALES PT"
        public DataTable setDataTablePT()
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[37] {
                new DataColumn("IdPagatodo", typeof(string)),
                new DataColumn("Fecha", typeof(string)),
                new DataColumn("Mes_Anio", typeof(string)),
                new DataColumn("Primeratransaccion", typeof(string)),
                new DataColumn("Hora", typeof(string)),
                new DataColumn("Tipo", typeof(string)),
                new DataColumn("IdSocio", typeof(string)),
                new DataColumn("Socio", typeof(string)),
                new DataColumn("IdCliente", typeof(string)),
                new DataColumn("Cliente", typeof(string)),
                new DataColumn("CorreoUsuario", typeof(string)),
                new DataColumn("FechaRegistro", typeof(string)),
                new DataColumn("Edad", typeof(string)),
                new DataColumn("Pais", typeof(string)),
                new DataColumn("Monto", typeof(decimal)),
                new DataColumn("ComisionPagatodo", typeof(decimal)),
                new DataColumn("ComisionPagaT", typeof(decimal)),
                new DataColumn("ComisionPalpal", typeof(decimal)),
                new DataColumn("Total", typeof(decimal)),
                new DataColumn("Resultado", typeof(string)),
                new DataColumn("idpaypal", typeof(string)),
                new DataColumn("IpRegistro", typeof(string)),
                new DataColumn("IpTransaccion", typeof(string)),
                new DataColumn("TotalTransacciones", typeof(string)),
                new DataColumn("Premovimientos", typeof(string)),
                new DataColumn("TotalIPs", typeof(string)),
                new DataColumn("totalcard", typeof(string)),
                new DataColumn("titular", typeof(string)),
                new DataColumn("envio_sms_recarga", typeof(string)),
                new DataColumn("certificado", typeof(string)),
                new DataColumn("destinatario", typeof(string)),
                new DataColumn("Comentarios", typeof(string)),
                new DataColumn("Modificaciones", typeof(string)),
                new DataColumn("Fraude", typeof(string)),
                new DataColumn("Comentarios_internos", typeof(string)),
                new DataColumn("Metodo_Pago", typeof(string)),
                new DataColumn("ComisionMasIVA", typeof(decimal)),
            });
            return dt;

        }
        public MemoryStream GetMemoryStreamPT(List<ReportPT> lista, DataTable dt)
        {
            decimal TotalAmount = 0;
            decimal TotalComisionPagatodo = 0;
            decimal TotalComisionPagaT = 0;
            decimal TotalComisionPalpal = 0;
            decimal Total = 0;
            decimal ComisionIVA = 0;
            foreach (var item1 in lista)
            {
                TotalAmount += item1.Monto;
                TotalComisionPagatodo += item1.ComisionPagatodo;
                TotalComisionPagaT += item1.ComisionPagaT;
                TotalComisionPalpal += item1.ComisionPalpal;
                Total += item1.Total;
                ComisionIVA += item1.ComisionMasIVA;
                dt.Rows.Add(
                    item1.IdPagatodo,
                    item1.Fecha,
                    item1.Mes_Anio,
                    item1.Primeratransaccion,
                    item1.Hora,
                    item1.Tipo,
                    item1.IdSocio,
                    item1.Socio,
                    item1.IdCliente,
                    item1.Cliente,
                    item1.CorreoUsuario,
                    item1.FechaRegistro,
                    item1.Edad,
                    item1.Pais,
                    item1.Monto,
                    item1.ComisionPagatodo,
                    item1.ComisionPagaT,
                    item1.ComisionPalpal,
                    item1.Total,
                    item1.Resultado,
                    item1.idpaypal,
                    item1.IpRegistro,
                    item1.IpTransaccion,
                    item1.TotalTransacciones,
                    item1.Premovimientos,
                    item1.TotalIPs,
                    item1.totalcard,
                    item1.titular,
                    item1.envio_sms_recarga,
                    item1.certificado,
                    item1.destinatario,
                    item1.Comentarios,
                    item1.Modificaciones,
                    item1.Fraude,
                    item1.Comentarios_internos,
                    item1.Metodo_Pago,
                    item1.ComisionMasIVA
                    );
            }
            dt.Rows.Add("-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "TOTAL", TotalAmount, TotalComisionPagatodo, TotalComisionPagaT, TotalComisionPalpal, Total, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", ComisionIVA);



            MemoryStream stream;
            using (XLWorkbook wb = new XLWorkbook())
            {
                var rs = wb.Worksheets.Add("ReportePT");

                //rs.Cell("A1").Value = "ReporteRocket";
                //var titlerange = rs.Range("A2:M2");
                //titlerange.Merge().Style.Font.SetBold().Font.FontSize = 13;
                //titlerange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                
                rs.Range("A2:AK2").Style.Font.SetBold();
                rs.Range("A2:AK2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rs.Cells("A2:AK2").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.PowderBlue;
             

                rs.Cell("A2").Value = "IdPagatodo";
                rs.Cell("B2").Value = "Fecha";
                rs.Cell("C2").Value = "Mes-Año";
                rs.Cell("D2").Value = "1°Transaccion";
                rs.Cell("E2").Value = "Hora";
                rs.Cell("F2").Value = "Tipo";
                rs.Cell("G2").Value = "IdSocio";
                rs.Cell("H2").Value = "Socio";
                rs.Cell("I2").Value = "IdCliente";
                rs.Cell("J2").Value = "Cliente";
                rs.Cell("K2").Value = "CorreoUsuario";
                rs.Cell("L2").Value = "FechaRegistro";
                rs.Cell("M2").Value = "Edad";
                rs.Cell("N2").Value = "Pais";
                rs.Cell("O2").Value = "Monto";
                rs.Cell("P2").Value = "ComisionPagatodo";
                rs.Cell("Q2").Value = "ComisionPagaT";
                rs.Cell("R2").Value = "ComisionPalpal";
                rs.Cell("S2").Value = "Total";
                rs.Cell("T2").Value = "Resultado";
                rs.Cell("U2").Value = "idpaypal";
                rs.Cell("V2").Value = "IpRegistro";
                rs.Cell("W2").Value = "IpTransaccion";
                rs.Cell("X2").Value = "TotalTransacciones";
                rs.Cell("Y2").Value = "Premovimientos";
                rs.Cell("Z2").Value = "TotalIPs";
                rs.Cell("AA2").Value = "totalcard";
                rs.Cell("AB2").Value = "titular";
                rs.Cell("AC2").Value = "envio_sms_recarga";
                rs.Cell("AD2").Value = "certificado";
                rs.Cell("AE2").Value = "destinatario";
                rs.Cell("AF2").Value = "Comentarios";
                rs.Cell("AG2").Value = "Modificaciones";
                rs.Cell("AH2").Value = "Fraude";
                rs.Cell("AI2").Value = "Comentarios_internos";
                rs.Cell("AJ2").Value = "Metodo_Pago";
                rs.Cell("AK2").Value = "ComisionMasIVA";

                rs.Cell("A3").InsertData(dt.Rows);

                rs.CellsUsed().Style.Border.BottomBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.BottomBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.TopBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.TopBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.LeftBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.LeftBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Border.RightBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rs.CellsUsed().Style.Border.RightBorderColor = ClosedXML.Excel.XLColor.Black;
                rs.CellsUsed().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rs.Columns().AdjustToContents();
                stream = GetStream(wb); //Método se define arriba


            }

            return stream;
        }
        #endregion

    }
}
