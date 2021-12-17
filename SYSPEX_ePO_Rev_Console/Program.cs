using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SYSPEX_ePO_Rev_Console
{

    class Program
    {
        #region ***** SQL Connection*****
        static readonly SqlConnection SGConnection = new SqlConnection("Server=192.168.1.21;Database=SYSPEX_LIVE;Uid=Sa;Pwd=Password1111;");
        static readonly SqlConnection JBConnection = new SqlConnection("Server=192.168.1.21;Database=Syspex Technologies (M) Sdn Bhd;Uid=Sa;Pwd=Password1111;");
        static SqlConnection SAPCon12 = new SqlConnection("Server=192.168.1.21;Database=AndriodAppDB;Uid=Sa;Pwd=Password1111;");
        static string SQLQuery;
        #endregion
        static void Main(string[] args)
        {
            EPO_REVISON("65ST");
            System.Threading.Thread.Sleep(6000);
            EPO_REVISON("07ST"); // go live 23/09/20
            System.Threading.Thread.Sleep(6000);
            EPO_REVISON("04SI"); // go live 25/09/20
            System.Threading.Thread.Sleep(6000);
            EPO_REVISON("03SM"); // go live 23/09/20
            System.Threading.Thread.Sleep(6000);

        }

        private static void EPO_REVISON(string companyCode)
        {
            bool flag;
            DataSet ds = GetPoRevNos(companyCode);

            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    flag = SendInvociePDF(ds.Tables[0].Rows[i]["docnum"].ToString(),
                                     ds.Tables[0].Rows[i]["docentry"].ToString(),
                                     ds.Tables[0].Rows[i]["E_Mail"].ToString(),
                                     ds.Tables[0].Rows[i]["cc"].ToString(), companyCode, ds.Tables[0].Rows[i]["cardname"].ToString(),
                                     ds.Tables[0].Rows[i]["U_RevNo"].ToString());

                    if (flag == true)
                    {
                        SqlCommand cmd = new SqlCommand();
                        SAPCon12.Open();
                        cmd.Connection = SAPCon12;
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "UPDATE syspex_ePO Set Revision = '" + ds.Tables[0].Rows[i]["U_RevNo"].ToString() + "'  where DocNum ='" + ds.Tables[0].Rows[i]["docnum"].ToString() + "' and Company ='" + companyCode + "'";
                        cmd.ExecuteNonQuery();
                        SAPCon12.Close();
                    }
                }
            }
        }

        private static DataSet GetPoRevNos(string companyCode)
        {
            SqlConnection SQLConnection = new SqlConnection();

            if (companyCode == "65ST")
                SQLConnection = SGConnection;
            SQLQuery = "Get_ePO_Revison";

            if (companyCode == "07ST")
                SQLConnection = JBConnection;
            SQLQuery = "Get_ePO_Revison";

            if (companyCode == "03SM")
                SQLConnection = JBConnection;
            SQLQuery = "Get_ePO_Revison";

            if (companyCode == "04SI")
                SQLConnection = JBConnection;
            SQLQuery = "Get_ePO_Revison";

            DataSet dsetItem = new DataSet();
            SqlCommand CmdItem = new SqlCommand(SQLQuery, SQLConnection)
            {
                CommandType = CommandType.StoredProcedure
            };
            SqlDataAdapter AdptItm = new SqlDataAdapter(CmdItem);
            AdptItm.Fill(dsetItem);
            CmdItem.Dispose();
            AdptItm.Dispose();
            SQLConnection.Close();
            return dsetItem;
        }


        private static bool SendInvociePDF(string DocNum, string DocEntry, string To, string CC, string CompanyCode, string VendorName, string Revision)
        {
            bool success;
            string Databasename = "";


            if (CompanyCode == "65ST")
                Databasename = "SYSPEX_LIVE";
            if (CompanyCode == "03SM")
                Databasename = "Syspex Mechatronic (M) Sdn Bhd";
            if (CompanyCode == "07ST")
                Databasename = "Syspex Technologies (M) Sdn Bhd";
            if (CompanyCode == "21SK")
                Databasename = "PT SYSPEX KEMASINDO";
            if (CompanyCode == "31SM")
                Databasename = "PT SYSPEX MULTITECH";
            if (CompanyCode == "04SI")
                Databasename = "Syspex Industries (M) Sdn Bhd";

            try
            {

                ReportDocument cryRpt = new ReportDocument();

                if ((CompanyCode == "03SM") || (CompanyCode == "04SI"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_03SM&04SI.rpt");

                if ((CompanyCode == "21SK") || (CompanyCode == "31SM"))
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_21SK&31SM.rpt");

                if (CompanyCode == "07ST")
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_07ST.rpt");

                if (CompanyCode == "65ST")
                    cryRpt.Load("F:\\Crystal Reports\\SYSPEX_PURCHASE_65ST.rpt");

                new TableLogOnInfos();
                TableLogOnInfo crtableLogoninfo;
                var crConnectionInfo = new ConnectionInfo();

                ParameterFieldDefinitions crParameterFieldDefinitions;
                ParameterFieldDefinition crParameterFieldDefinition;
                ParameterValues crParameterValues = new ParameterValues();
                ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

                crParameterDiscreteValue.Value = Convert.ToString(DocEntry);
                crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
                crParameterFieldDefinition = crParameterFieldDefinitions["@DOCENTRY"];
                crParameterValues = crParameterFieldDefinition.CurrentValues;

                crParameterValues.Clear();
                crParameterValues.Add(crParameterDiscreteValue);
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

                crConnectionInfo.ServerName = "SYSPEXSAP04";
                crConnectionInfo.DatabaseName = Databasename;
                crConnectionInfo.UserID = "sa";
                crConnectionInfo.Password = "Password1111";

                var crTables = cryRpt.Database.Tables;
                foreach (Table crTable in crTables)
                {
                    crtableLogoninfo = crTable.LogOnInfo;
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                    crTable.ApplyLogOnInfo(crtableLogoninfo);
                }



                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();

                CrDiskFileDestinationOptions.DiskFileName = "F:\\ePORev\\" + CompanyCode + "\\" + DocNum + ".pdf";
                CrExportOptions = cryRpt.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                cryRpt.Export();

                //// Email Part 

                MailMessage mm = new MailMessage
                {
                    From = new MailAddress("noreply@syspex.com")
                };


                //CC Address
                foreach (var address in CC.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Distinct())
                {
                    mm.CC.Add(new MailAddress(address)); //Adding Multiple CC email Id
                }

                //TO Address

                foreach (var address in To.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (IsValidEmail(address) == true)
                    {
                        mm.To.Add(address);
                    }
                }

                mm.IsBodyHtml = true;
                mm.Subject = "Amended PO#" + DocNum + " " + Revision + "_" + VendorName;

                if (CompanyCode != "65ST")
                {
                    //mm.Subject = " Purchase Order No:" + DocNum;
                    mm.Body = "<p>Dear Valued Supplier,</p> <p>Attached please find our <u>PO# " + DocNum + "(" + Revision + ")</u>, if you have any questions please call us immediately.</p>" +
                        "<p> Regards,</p>" +
          "<p> Procurement Team</p> ";


                }
                else
                {
                    mm.Body = ST_HTMLBULIDER(DocNum, Revision);
                }

                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    EnableSsl = true
                };
                if (CompanyCode == "65ST")
                {
                    System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential("sg.procurement@syspex.com", "enhance5");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                    smtp.Send(mm);
                }
                else
                {
                    System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential("noreply@syspex.com", "design360");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    mm.Attachments.Add(new System.Net.Mail.Attachment(CrDiskFileDestinationOptions.DiskFileName));
                    smtp.Send(mm);
                }
                success = true;


            }
            catch (CrystalReportsException ex)
            {

                throw ex;
            }

            return success;
        }
        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;

            }
            catch
            {
                return false;
            }
        }
        private static string ST_HTMLBULIDER(string DocNum, string Revision)
        {

            //Create a new StringBuilder object
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<p>Dear Supplier,</p>");
            sb.AppendLine("<p>Please find <strong><u>PO# " + DocNum + "(" + Revision + ")</u></strong> and file attachments.</p>");
            sb.AppendLine("<p>Reply back this email to confirm on the order quantity and the delivery date stated on the PO within the next 24 hours</p>");
            sb.AppendLine("<p>Kindly take note and comply with the following packaging and delivery information, </p>");
            sb.AppendLine("<ol>");
            sb.AppendLine("<li>To indicate Syspex PO number for both Invoice and DO.</li>");
            sb.AppendLine("<li>To indicate serial number on each outer packaging (When applicable).</li>");
            sb.AppendLine("<li> To take note our receiving hours (Monday to Fridays 10:00am &ndash; 12:00 &amp; 1:00pm &ndash; 4:00pm).<strong>- Only applicable to supplier(s) deliver at Syspex Warehouse</strong></li>");
            sb.AppendLine("<li> Please take note and comply that total height of incoming palletised goods should not exceed 1.5m.</ li>");
            sb.AppendLine("<li> The pallet must be able to truck by hand pallet truck.</li>");
            sb.AppendLine("<li> Please email us soft copy of invoice and packing list once shipment ready for dispatch.</li>");
            sb.AppendLine("<li> For multiple package shipment, please indicate content list on outside of each package.</li>");
            sb.AppendLine("</ol>");
            sb.AppendLine("<p>Thank you for your co-operation.</p>");
            sb.AppendLine("<p>Best Regards,</p>");
            sb.AppendLine("<p>Syspex Procurement Team</p>");
            return sb.ToString();


        }
    }
}
