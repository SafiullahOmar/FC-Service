using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace FCWindowService
{
    public partial class FCScheduler : ServiceBase
    {
        private Timer timer1 = null;
        public FCScheduler()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            library ob = new library();
            ob.ExecuteQuery("insert into tbl_Region values(12,'khan','','','')");
            //timer1 = new Timer();
            //this.timer1.Interval = 10000;
            //this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            //timer1.Enabled = true;
        }

        public void func() {
            library ob = new library();
            ob.ExecuteQuery("insert into tbl_Region values(12,'khan','','','')");
        }
        public  void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            library db = new library();
            try
            {
                DataTable dt = db.SelectRecords(@";with tbl as
                                            (
                                                    SELECT        HC_FarmerActivityDetails.FarmerId, ISNULL(SUM(FC_FarmerContActivityDetails.Area), 0) AS area, ISNULL(SUM(FC_FarmerContActivityDetails.Amount), 0) AS TotalAmount, HC_FarmerActivityDetails.Year,
                                                                                 (SELECT        ISNULL(SUM(RecievedAmount), 0) AS RecievedAmount
                                                                                   FROM            FC_FarmerRecieveAmount
                                                                                   WHERE     cast( FC_FarmerRecieveAmount.EntryDate as datetime) between DateAdd(DD,-7,GETDATE()) and GETDATE()  and (FarmerId = HC_FarmerActivityDetails.FarmerId) AND (Year = HC_FarmerActivityDetails.Year)) AS RecievedAmount,
							                                                       ISNULL(SUM(FC_FarmerContActivityDetails.Amount), 0)-
							                                                       (SELECT        ISNULL(SUM(RecievedAmount), 0) 
                                                                                   FROM            FC_FarmerRecieveAmount
                                                                                   WHERE     cast( FC_FarmerRecieveAmount.EntryDate as datetime) between DateAdd(DD,-7,GETDATE()) and GETDATE()  and (FarmerId = HC_FarmerActivityDetails.FarmerId) AND (Year = HC_FarmerActivityDetails.Year))  AS Balance
							                                                        , OCM_Province.ProvinceEngName, OCM_District.DistrictEngName,FC_ExtensionWorkerInfo.Name
                                                    FROM            FC_ExtensionWorkerInfo INNER JOIN
                                                                             FC_FarmerInfo ON FC_ExtensionWorkerInfo.ExtWID = FC_FarmerInfo.ExtWId INNER JOIN
                                                                             FC_Season INNER JOIN
                                                                             HC_FarmerActivityDetails INNER JOIN
                                                                             FC_FarmerContActivityDetails ON HC_FarmerActivityDetails.FarmerActivityId = FC_FarmerContActivityDetails.FarmerActivityId ON FC_Season.SeasonId = HC_FarmerActivityDetails.SeasonId ON 
                                                                             FC_FarmerInfo.Id = HC_FarmerActivityDetails.FarmerId INNER JOIN
                                                                             OCM_Region INNER JOIN
                                                                             OCM_Province ON OCM_Region.id = OCM_Province.Region ON FC_ExtensionWorkerInfo.ProvinceID = OCM_Province.ProvinceID INNER JOIN
                                                                             OCM_District ON FC_ExtensionWorkerInfo.DistrictID = OCM_District.DistrictID
                                                    WHERE        (FC_FarmerContActivityDetails.IsDeleted = 0) AND (HC_FarmerActivityDetails.Year = 2017)
                                                    GROUP BY HC_FarmerActivityDetails.Year, HC_FarmerActivityDetails.FarmerId, OCM_Province.ProvinceEngName, OCM_District.DistrictEngName, FC_ExtensionWorkerInfo.Name
                                                    )
                                                    select  DateAdd(DD,-7,GETDATE()) as reportdate,ProvinceEngName as [Province Name],DistrictEngName as [District Name],Name as [Extension Worker],isnull(sum(TotalAmount),0) as [Collectable Amount],Year,isnull(sum(RecievedAmount),0) [Paid Amount],sum(Balance) as [Balance] from tbl
                                                    where RecievedAmount>0
                                                    group by Year,ProvinceEngName,DistrictEngName,Name

                                            ");
                DataTable dtNotPaid = db.SelectRecords(@";with tbl as
(

SELECT        HC_FarmerActivityDetails.FarmerId, ISNULL(SUM(FC_FarmerContActivityDetails.Area), 0) AS area, ISNULL(SUM(FC_FarmerContActivityDetails.Amount), 0) AS TotalAmount, HC_FarmerActivityDetails.Year,
                             (SELECT        ISNULL(SUM(RecievedAmount), 0) AS RecievedAmount
                               FROM            FC_FarmerRecieveAmount
                               WHERE     cast( FC_FarmerRecieveAmount.EntryDate as datetime) between DateAdd(DD,-7,GETDATE()) and GETDATE()  and (FarmerId = HC_FarmerActivityDetails.FarmerId) AND (Year = HC_FarmerActivityDetails.Year)) AS RecievedAmount,
							   ISNULL(SUM(FC_FarmerContActivityDetails.Amount), 0)-
							   (SELECT        ISNULL(SUM(RecievedAmount), 0) 
                               FROM            FC_FarmerRecieveAmount
                               WHERE     cast( FC_FarmerRecieveAmount.EntryDate as datetime) between DateAdd(DD,-7,GETDATE()) and GETDATE()  and (FarmerId = HC_FarmerActivityDetails.FarmerId) AND (Year = HC_FarmerActivityDetails.Year))  AS Balance
							    , OCM_Province.ProvinceEngName, OCM_District.DistrictEngName,FC_ExtensionWorkerInfo.Name
FROM            FC_ExtensionWorkerInfo INNER JOIN
                         FC_FarmerInfo ON FC_ExtensionWorkerInfo.ExtWID = FC_FarmerInfo.ExtWId INNER JOIN
                         FC_Season INNER JOIN
                         HC_FarmerActivityDetails INNER JOIN
                         FC_FarmerContActivityDetails ON HC_FarmerActivityDetails.FarmerActivityId = FC_FarmerContActivityDetails.FarmerActivityId ON FC_Season.SeasonId = HC_FarmerActivityDetails.SeasonId ON 
                         FC_FarmerInfo.Id = HC_FarmerActivityDetails.FarmerId INNER JOIN
                         OCM_Region INNER JOIN
                         OCM_Province ON OCM_Region.id = OCM_Province.Region ON FC_ExtensionWorkerInfo.ProvinceID = OCM_Province.ProvinceID INNER JOIN
                         OCM_District ON FC_ExtensionWorkerInfo.DistrictID = OCM_District.DistrictID
WHERE        (FC_FarmerContActivityDetails.IsDeleted = 0) AND (HC_FarmerActivityDetails.Year = 2017)
GROUP BY HC_FarmerActivityDetails.Year, HC_FarmerActivityDetails.FarmerId, OCM_Province.ProvinceEngName, OCM_District.DistrictEngName, FC_ExtensionWorkerInfo.Name
)

select  DateAdd(DD,-7,GETDATE()) as reportdate,ProvinceEngName as [Province Name],DistrictEngName as [District Name],Name as [Extension Worker],isnull(sum(TotalAmount),0) as [Collectable Amount],Year,isnull(sum(RecievedAmount),0) [Paid Amount],sum(Balance) as [Balance] from tbl
where RecievedAmount=0
group by Year,ProvinceEngName,DistrictEngName,Name

                                            ");
                DataTable dtEmail = db.SelectRecords(@"select Email from tbl_PC 
where Email is not null
union all
select email from tbl_Region
where Email is not null
union all
select email from aspnet_Membership where ApplicationId like '34764d90-4054-47ea-ab95-17c5d32a6136' and email is not null");

                if (dt.Rows.Count > 0 || dtNotPaid.Rows.Count > 0)
                {
                    byte[] bytes, bytesNotPaid;

                    string emails = "";
                    if (dtEmail.Rows.Count > 0)
                    {
                        foreach (DataRow rw in dtEmail.Rows)
                        {
                            emails += rw["Email"].ToString() + ",";
                        }
                    }
                    emails += "omar.safiullah@gmail.com,grkinyanjui@gmail.com,khalid.ferdaus@yahoo.com,usman.safi@mail.gov.af,arshidir@live.com,h_saleemsafi@yahoo.com,samadsame@gmail.com,khalid.ibrahimsafi@gmail.com,shakir_vet@yahoo.com,shaimaahadi2003@yahoo.com";
                    MailMessage mm = new MailMessage("fcmis.nhlp@gmail.com", "omar.safiullah@gmail.com");
                    mm.Subject = "eWeeklyReport:Farmer Contribution Report for NHLP/MAIL";
                    string body = "<strong>Dear Observers,</strong>  Thank you for seeing online report of  <strong>Farmer Contribution Managment Information System</strong>. Please find the attached reports which is generated from date :<strong>" + dtNotPaid.Rows[0]["reportdate"].ToString() + " up to date :" + DateTime.Now.ToString() + " </strong> in seperate files , extension workers who paid and those who havent paid";
                    body += "<br/>If you are unable to read the PDF file, please download the latest version of the Adobe Acrobat Reader:<br/>http://get.adobe.com/reader/</br>";
                    body += "<br/>For further information or queries, please contact<strong> NHLP MIS Department</strong>, or go for online on <strong>http:// 103.13.66.210:88</strong> ";
                    body += "<br/>------------------------------------------------------------------------------<br/><br/>This is an automatically generated message; please do not reply to this email.";
                    mm.Body = body;
                    if (dt.Rows.Count > 0)
                    {
                        bytes = SendPDFEmail(dt);
                        mm.Attachments.Add(new Attachment(new MemoryStream(bytes), "FCMIS" + DateTime.Now.ToShortDateString() + "PaidAmountreport.pdf"));
                    }

                    if (dtNotPaid.Rows.Count > 0)
                    {
                        bytesNotPaid = SendPDFNotPaidEmail(dtNotPaid);
                        mm.Attachments.Add(new Attachment(new MemoryStream(bytesNotPaid), "FCMIS" + DateTime.Now.ToShortDateString() + "NotPaidAmountreport.pdf"));
                    }

                    mm.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    NetworkCredential NetworkCred = new NetworkCredential();
                    NetworkCred.UserName = "fcmis.nhlp@gmail.com";
                    NetworkCred.Password = "safi_khan123";
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = 587;
                    smtp.Send(mm);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                db.Connection.Close();
            }
        }
        protected override void OnStop()
        {
            timer1.Enabled = false;
        }

        private byte[] SendPDFEmail(DataTable dt)
        {
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2'>");
                    sb.Append("<tr><td align='center' ></td align='center' ><td align='center'>Farmer Contribution Managment Information System</td><td align='center' ></td></tr>");
                    sb.Append("<tr><td align='center' style='background-color: #18B5F0' colspan = '3'><b>NHLP Farmer Contribution Report</b></td></tr>");
                    sb.Append("<tr><td colspan = '3'></td></tr>");
                    sb.Append("<tr><td><b>Farmer Contribution Detail By Extension Worker , District and Province</b>");

                    sb.Append("</td><td colspan = '2'><b>Report Date: </b>");
                    sb.Append("From Date :" + dt.Rows[0]["reportdate"].ToString() + " Up To : " + DateTime.Now);

                    sb.Append(" </td></tr>");

                    sb.Append("<tr><td colspan = '3'><b>List of Extension Workers By districts who Paid within mentioned period :</b></td></tr>");
                    sb.Append("</table>");
                    sb.Append("<br />");
                    sb.Append("<table border = '1'>");
                    sb.Append("<tr>");
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName != "reportdate")
                        {
                            sb.Append("<td bgcolor='orange'><font color='white'>");
                            sb.Append("<strong>" + column.ColumnName + "</strong>");
                            sb.Append("</font></td>");
                        }
                    }
                    sb.Append("</tr>");
                    string pName = null;
                    double collectable = 0, paid = 0, balnce = 0;
                    int rowslenght = dt.Rows.Count;
                    foreach (DataRow row in dt.Rows)
                    {
                        rowslenght--;
                        if (pName == null)
                            pName = row["Province Name"].ToString();
                        else
                        {
                            if (pName != row["Province Name"].ToString())
                            {

                                sb.Append("<tr bgcolor='black' color='white' ><td colspan = '3'><b>Summary for " + pName.ToString() + " Province </b></td><td><b>" + collectable.ToString() + "</b></td><td><b>Paid Amount</b></td><td><b>" + paid.ToString() + "</b></td>");

                                sb.Append("<td>" + balnce.ToString() + "</td></tr>");
                                pName = row["Province Name"].ToString();
                                collectable = 0; paid = 0; balnce = 0;
                            }
                        }

                        collectable += Convert.ToDouble(row["Collectable Amount"].ToString());
                        paid += Convert.ToDouble(row["Paid Amount"].ToString());
                        balnce += Convert.ToDouble(row["Balance"].ToString());
                        sb.Append("<tr>");
                        foreach (DataColumn column in dt.Columns)
                        {
                            if (column.ColumnName != "reportdate")
                            {
                                sb.Append("<td>");
                                sb.Append(row[column]);
                                sb.Append("</td>");
                            }
                        }
                        sb.Append("</tr>");

                        if (rowslenght == 0)
                        {

                            sb.Append("<tr bgcolor='black' color='white' ><td colspan = '3'><b>Summary for " + pName.ToString() + " Province </b></td><td><b>" + collectable.ToString() + "</b></td><td><b>Paid Amount</b></td><td><b>" + paid.ToString() + "</b></td>");

                            sb.Append("<td>" + balnce.ToString() + "</td></tr>");
                            pName = row["Province Name"].ToString();
                        }
                    }
                    sb.Append("</table>");
                    StringReader sr = new StringReader(sb.ToString());

                    Document pdfDoc = new Document(iTextSharp.text.PageSize.A4.Rotate());
                    HTMLWorker htmlparser = new HTMLWorker(pdfDoc);

                    // Net

                    //1st Pdf
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                        pdfDoc.Open();
                        htmlparser.Parse(sr);
                        pdfDoc.Close();
                        byte[] bytes = memoryStream.ToArray();
                        memoryStream.Close();
                        return bytes;

                    }
                }
            }
        }
        private byte[] SendPDFNotPaidEmail(DataTable dt)
        {
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("<table width='100%' cellspacing='0' cellpadding='2'>");
                    sb.Append("<tr><td align='center' ></td align='center' ><td align='center'>Farmer Contribution Managment Information System</td><td align='center' ></td></tr>");
                    sb.Append("<tr><td align='center' style='background-color: #18B5F0' colspan = '3'><b>NHLP Farmer Contribution Report</b></td></tr>");
                    sb.Append("<tr><td colspan = '3'></td></tr>");
                    sb.Append("<tr><td><b>Farmer Contribution Detail By Extension Worker , District and Province</b>");

                    sb.Append("</td><td colspan = '2'><b>Report Date: </b>");
                    sb.Append("From Date :" + dt.Rows[0]["reportdate"].ToString() + " Up To : " + DateTime.Now);

                    sb.Append(" </td></tr>");

                    sb.Append("<tr><td colspan = '3'><b>List of Extension Workers By districts who haven't Paid within mentioned period :</b></td></tr>");
                    sb.Append("</table>");
                    sb.Append("<br />");
                    sb.Append("<table border = '1'>");
                    sb.Append("<tr>");
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName != "reportdate")
                        {
                            sb.Append("<td bgcolor='red'><font color='white'>");
                            sb.Append("<strong>" + column.ColumnName + "</strong>");
                            sb.Append("</font></td>");
                        }
                    }
                    sb.Append("</tr>");
                    string pName = null;
                    double collectable = 0, paid = 0, balnce = 0;
                    int rowslenght = dt.Rows.Count;
                    foreach (DataRow row in dt.Rows)
                    {
                        rowslenght--;
                        if (pName == null)
                            pName = row["Province Name"].ToString();
                        else
                        {
                            if (pName != row["Province Name"].ToString())
                            {

                                sb.Append("<tr bgcolor='gray' color='white' ><td colspan = '3'><b>Summary for " + pName.ToString() + " Province </b></td><td><b>" + collectable.ToString() + "</b></td><td><b>Paid Amount</b></td><td><b>" + paid.ToString() + "</b></td>");

                                sb.Append("<td>" + balnce.ToString() + "</td></tr>");
                                pName = row["Province Name"].ToString();
                                collectable = 0; paid = 0; balnce = 0;
                            }
                        }

                        collectable += Convert.ToDouble(row["Collectable Amount"].ToString());
                        paid += Convert.ToDouble(row["Paid Amount"].ToString());
                        balnce += Convert.ToDouble(row["Balance"].ToString());
                        sb.Append("<tr>");
                        foreach (DataColumn column in dt.Columns)
                        {
                            if (column.ColumnName != "reportdate")
                            {
                                sb.Append("<td>");
                                sb.Append(row[column]);
                                sb.Append("</td>");
                            }
                        }
                        sb.Append("</tr>");

                        if (rowslenght == 0)
                        {

                            sb.Append("<tr bgcolor='gray' color='white' ><td colspan = '3'><b>Summary for " + pName.ToString() + " Province </b></td><td><b>" + collectable.ToString() + "</b></td><td><b>Paid Amount</b></td><td><b>" + paid.ToString() + "</b></td>");

                            sb.Append("<td>" + balnce.ToString() + "</td></tr>");
                            pName = row["Province Name"].ToString();
                        }
                    }
                    sb.Append("</table>");
                    StringReader sr = new StringReader(sb.ToString());

                    Document pdfDoc = new Document(iTextSharp.text.PageSize.A4.Rotate());
                    HTMLWorker htmlparser = new HTMLWorker(pdfDoc);

                    // Net

                    //1st Pdf
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        PdfWriter writer = PdfWriter.GetInstance(pdfDoc, memoryStream);
                        pdfDoc.Open();
                        htmlparser.Parse(sr);
                        pdfDoc.Close();
                        byte[] bytes = memoryStream.ToArray();
                        memoryStream.Close();
                        return bytes;

                    }
                }
            }
        }

    }
}
