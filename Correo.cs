using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.IO;


namespace ReportePTGPS
{
   public class Correo
    {
        public void SMTP4_1(Stream file, Stream file2,Stream file3, Stream file4, Stream file5, Stream file6)
        {
            string todate = getDate(-3);
            string today2 = getDate(-2);
            string today3 = getDate(-1);

            Enviar4_1(
                file,
                file2,
                file3,
                file4, 
                file5,
                file6, 
                "Reporte PagaTodo",
                "ReportePTRocket-" + todate + ".xlsx",
                "ReportePTRocket-" + today2 + ".xlsx",
                "ReportePTRocket-" + today3 + ".xlsx",
                "ReportePT-" + todate + ".xlsx",
                "ReportePT-" + today2 + ".xlsx",
                "ReportePT-" + today3 + ".xlsx"
                );
        }

        public void SMTP4_2(Stream file, Stream file2)
        {
            string todate = DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy");

            Enviar4_2(
                file,
                file2,
                "Reporte PagaTodo",
                "ReportePTRocket-" + todate + ".xlsx",
                "ReportePT-" + todate + ".xlsx"
                );
        }
        public string Enviar4_1(
            Stream file, 
            Stream file2,
            Stream file3, 
            Stream file4, 
            Stream file5,
            Stream file6, 
            string subject,
            string filename,
            string filename2,
            string filename3,
            string filename4,
            string filename5,
            string filename6
            )
        {
            string result = "";
            MailMessage email = new MailMessage();

            Attachment data = new Attachment(file, filename);
            FileInfo fileinfo = new FileInfo(filename);
            string directory = fileinfo.Directory.Parent.FullName;

            Attachment data2 = new Attachment(file2, filename2);
            FileInfo fileinfo2 = new FileInfo(filename2);
            string directory2 = fileinfo.Directory.Parent.FullName;

            Attachment data3 = new Attachment(file3, filename3);
            FileInfo fileinfo3 = new FileInfo(filename3);
            string directory3 = fileinfo.Directory.Parent.FullName;

            Attachment data4 = new Attachment(file4, filename4);
            FileInfo fileinfo4 = new FileInfo(filename4);
            string directory4 = fileinfo.Directory.Parent.FullName;

            Attachment data5 = new Attachment(file5, filename5);
            FileInfo fileinfo5 = new FileInfo(filename5);
            string directory5 = fileinfo.Directory.Parent.FullName;

            Attachment data6 = new Attachment(file6, filename6);
            FileInfo fileinfo6 = new FileInfo(filename6);
            string directory6 = fileinfo.Directory.Parent.FullName;

            string Fridaydate = DateTime.Now.AddDays(-3).ToString("dd-MM-yyyy");
            string FridayDay = DateTime.UtcNow.AddDays(-3).ToString("dddd");

            string SaturdayDate = DateTime.Now.AddDays(-2).ToString("dd-MM-yyyy");
            string SaturdayDay = DateTime.UtcNow.AddDays(-2).ToString("dddd");

            string SundayDate = DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy");
            string SundayDay = DateTime.UtcNow.AddDays(-1).ToString("dddd");

            email.To.Add(new MailAddress("eescobar@globalpay.us"));
            email.To.Add(new MailAddress("blopez@globalpay.us"));
            email.To.Add(new MailAddress("ehernandez@globalpay.us"));
            email.To.Add(new MailAddress("federicobp@globalpay.us"));
            email.To.Add(new MailAddress("victor@globalpay.us"));
            email.To.Add(new MailAddress("gescobar@globalpay.us"));
            email.To.Add(new MailAddress("aizaguirre@globalpay.us"));
            email.To.Add(new MailAddress("nmonge@globalpay.us"));
            email.To.Add(new MailAddress("cbolanos@globalpay.us"));





            email.From = new MailAddress("noreply@globalpay.us");
            email.Subject = subject;
            emai l.Attachments.Add(data);
            email.Attachments.Add(data2);
            email.Attachments.Add(data3);
            email.Attachments.Add(data4);
            email.Attachments.Add(data5);
            email.Attachments.Add(data6);

            email.Body = "¡Buen día!. Envío Los Reportes de Pagos: \n     " + FridayDay +": "+ Fridaydate+ ", \n     " + SaturdayDay + ": "+ SaturdayDate+",\n     "+SundayDay+": "+SundayDate+ "\n¡Bendiciones!";
            email.IsBodyHtml = false;
            email.Priority = MailPriority.Normal;
             
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "mail.globalpay.us";
            smtp.Port = 587;
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("noreply@globalpay.us", "n0r3ply2021-");

            try
            {
                smtp.Send(email);
                email.Dispose();
                result = "exito";
            }
            catch (Exception ex)
            {
                result = "error";
            }
            Console.WriteLine(result);
            return result;
        }
        public string Enviar4_2(
            Stream file,
            Stream file2,
            string subject, 
            string filename,
            string filename2)
        {
            string result = "";
            MailMessage email = new MailMessage();

            Attachment data = new Attachment(file, filename);
            FileInfo fileinfo = new FileInfo(filename);
            string directory = fileinfo.Directory.Parent.FullName;

            Attachment data2 = new Attachment(file2, filename2);
            FileInfo fileinfo2 = new FileInfo(filename2);
            string directory2 = fileinfo.Directory.Parent.FullName;


            string yesterday = DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy");
            string yesterdayDay = DateTime.UtcNow.AddDays(-1).ToString("dddd");


            email.To.Add(new MailAddress("eescobar@globalpay.us"));
            email.To.Add(new MailAddress("blopez@globalpay.us"));
            email.To.Add(new MailAddress("ehernandez@globalpay.us"));
            email.To.Add(new MailAddress("federicobp@globalpay.us"));
            email.To.Add(new MailAddress("victor@globalpay.us"));
            email.To.Add(new MailAddress("gescobar@globalpay.us"));
            email.To.Add(new MailAddress("aizaguirre@globalpay.us"));
            email.To.Add(new MailAddress("nmonge@globalpay.us"));
            email.To.Add(new MailAddress("cbolanos@globalpay.us"));


            email.From = new MailAddress("noreply@globalpay.us");
            email.Subject = subject;
            email.Attachments.Add(data);
            email.Attachments.Add(data2);
            email.Body = "¡Buen día!. Envío Los Reportes de Pagos: \n     " + yesterdayDay + ": " + yesterday + "\n¡Bendiciones!";
            email.IsBodyHtml = false;
            email.Priority = MailPriority.Normal;

            SmtpClient smtp = new SmtpClient();
            smtp.Host = "mail.globalpay.us";
            smtp.Port = 587;
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("noreply@globalpay.us", "n0r3ply2021-");

            try
            {
                smtp.Send(email);
                email.Dispose();
                result = "exito";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
          
            return result;
        }

        //Con esta funcion obtenemos el dia que deseamos, solo le pasamos los dias que queremos
        //sumar o restar ejemplo getDate(-2) o getDate(2)
        public string getDate(int day)
        {
            string date = DateTime.Now.AddDays(day).ToString("dd-MM-yyyy");
            return date;
        }
    }
}
