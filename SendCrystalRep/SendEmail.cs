using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Net.Mail;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Specialized;
using System.Windows.Forms;
using SAPbobsCOM;
using System.Net;
using System.Text;
using System.Web;

namespace SendCrystalRep
{

    public partial class SendEmail : Form
    {

        class Parameter
        {
            public string ParameterName { get; set; }
            public string ParameterValue { get; set; }
        }

        SAPbobsCOM.Company oCompany;
        string mail;
        string mailPassword;
        string smpt;
        String pathRep;
        String ReportPath;
        ReportDocument reportDocument;
        String ReportKey;
        String ReportValue;
        List<Parameter> ParameterValues = new List<Parameter>();
        String pdfFile;
       // string Val;

        public SendEmail()
        {
            oCompany = new SAPbobsCOM.Company();

            //Database connection
            String sqlServer = System.Configuration.ConfigurationManager.AppSettings["SS"];
            if (sqlServer == "SQL2012") oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
            else if (sqlServer == "SQL2014") oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            else if (sqlServer == "SQL2016") oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            else if (sqlServer == "HANA") oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            else oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL;

            oCompany.Server = ConfigurationManager.AppSettings["DS"];
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = Convert.ToBoolean(ConfigurationManager.AppSettings["AU"]);
            oCompany.DbUserName = ConfigurationManager.AppSettings["UN"];
            oCompany.DbPassword = ConfigurationManager.AppSettings["PW"];
            oCompany.CompanyDB = ConfigurationManager.AppSettings["DB"];
            oCompany.UserName = ConfigurationManager.AppSettings["SAPUN"];
            oCompany.Password = ConfigurationManager.AppSettings["SAPPW"];
            mail = ConfigurationManager.AppSettings["email"];
            mailPassword = ConfigurationManager.AppSettings["ePw"];
            smpt = ConfigurationManager.AppSettings["smpt"];
            ReportPath = ConfigurationManager.AppSettings["path"];
            ReportKey = ConfigurationManager.AppSettings["RepKey"];
            ReportValue = ConfigurationManager.AppSettings["val"];

            int con = oCompany.Connect();
            string error = oCompany.GetLastErrorDescription();
            //End
            InitializeComponent();

            //Get report
            pathRep = ReportPath;

            string key = ReportKey;
            string val = ReportValue;
            ;

            Parameter param = new Parameter();
                param.ParameterName = key;
                val = val.Replace("%20", " ");
                param.ParameterValue = val;
                ParameterValues.Add(param);
        }

        private void SendEmail_Load(object sender, EventArgs e)
        {

        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                //Report convert to PDF
                reportDocument = new ReportDocument();
                reportDocument.Load(pathRep);
                foreach (var item in ParameterValues)
                {
                    reportDocument.SetParameterValue(item.ParameterName, item.ParameterValue);
                }
        
                pdfFile = Path.GetTempPath() + "\\" + Guid.NewGuid() + ".pdf";
                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                CrDiskFileDestinationOptions.DiskFileName = pdfFile;
                CrExportOptions = reportDocument.ExportOptions;

                Object CrFormatTypeOptions;
                ExportFormatType eft;
                String ContentType;

                CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                eft = ExportFormatType.PortableDocFormat;

                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = eft;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }

                reportDocument.Export();
                //End
            }

            catch (Exception ex)
            {
                throw ex;
            }
            //Email
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("Select \"U_Suject\",\"U_Body\" from \"@AUTO_EMAIL\" ");

            string sa = "ReciverEmail@gmail.com";
            MailMessage message = new MailMessage();
            message.From = new MailAddress(mail);
            message.To.Add(sa);

            SAPbobsCOM.Recordset oRecordset1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("Select \"U_Suject\",\"U_Body\" from \"@AUTO_EMAIL\" ");

            String subject = oRecordset.Fields.Item("U_Suject").Value;
            String body = oRecordset.Fields.Item("U_Body").Value;

            message.Subject = subject;
            message.Body = body;

            String filePath = pdfFile;
            if (File.Exists(filePath))
            {
                message.Attachments.Add(new System.Net.Mail.Attachment(filePath));
            }

            message.BodyEncoding = Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient client = new SmtpClient(smpt, 587);
            System.Net.NetworkCredential basicCredential1 = new
            System.Net.NetworkCredential(mail, mailPassword);
            client.EnableSsl = true;
            client.UseDefaultCredentials = false;
            client.Credentials = basicCredential1;

            try
            {
                client.Send(message);
            }

            catch (Exception ex)
            {
                throw ex;
            }
 
   
            MessageBox.Show("OK");
            //End
        }
       
    }

   
}

