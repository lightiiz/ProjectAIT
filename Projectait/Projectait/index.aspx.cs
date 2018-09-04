using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Configuration;
using System.Data.OleDb;


namespace Projectait
{
    public partial class index : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;

        SqlConnection cn = new SqlConnection();
        SqlCommand sqlQuery = new SqlCommand();
        DataTable dt = new DataTable();
        string sql = "";

        SqlConnection cnn = new SqlConnection();
        SqlCommand sqlQueryn = new SqlCommand();
        DataTable dtn = new DataTable();
        string sqln = "";
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void bindgrid()
        {


            String txtDate1 = TextBox1.Text;
            String txtDate2 = TextBox2.Text;

            sql = "select [Circuit ID] as หมายเลขวงจร ";
            sql = sql + ",COALESCE(NULLIF(RIGHT([Subject],LEN([Subject])-charindex(')',[Subject])) , ''),SUBSTRING([Subject],1,case when charindex('(',[Subject]) = 0 then LEN([Subject]) else charindex('(',[Subject])-1 END)) as หน่วยงานผู้ใช้ ";
            sql = sql + ",[SLA]*100  as SLA ";
            sql = sql + ",ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103))  as วันที่แจ้งเหตุ ";
            sql = sql + ",ISNULL(CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) as เวลาแจ้งเหตุ ";
            sql = sql + ",[Help Topic] as ประเภทเหตุขัดข้อง ";
            sql = sql + ",SUBSTRING([สรุปสาเหตุ # วิธีแก้ปัญหา],1,case when charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา]) = 0 then LEN([สรุปสาเหตุ # วิธีแก้ปัญหา]) else charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])-1 END) as สาเหตุ ";
            sql = sql + ",RIGHT([สรุปสาเหตุ # วิธีแก้ปัญหา],LEN([สรุปสาเหตุ # วิธีแก้ปัญหา])-charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])) as การแก้ไข ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],103),CONVERT(VARCHAR(10),[Closed Date],103)) as วันที่แก้ไข ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) as เวลาแก้ไข ";
            sql = sql + ",(select CONVERT(VARCHAR(10),Datediff(ss,a.t1,a.t2)/(60*60*24)) + ' Days ' ";
            sql = sql + "   +CONVERT(VARCHAR(5),DateAdd(SS,Datediff(ss,a.t1, a.t2)%(60*60*24),0),114) ";
            sql = sql + "from ";
            sql = sql + "(select ISNULL (CONVERT(VARCHAR(10),[Link Down],102),CONVERT(VARCHAR(10),[Create Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) t1 ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],102),CONVERT(VARCHAR(10),[Closed Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) t2 ";
            sql = sql + ") a) as 'ระยะเวลาการแก้ไข' ";
            sql = sql + ",[Ticket Number] as 'OS Ticket Number' ";
            sql = sql + ",CONVERT(VARCHAR(3),DATENAME(dw, ISNULL ([Link Down],[Create Date]))) +'-'+CONVERT(VARCHAR(3),DATENAME(dw, ISNULL ([Link Up],[Closed Date]))) as 'หมายเหตุ' ";
            sql = sql + ",[Root Cause] as 'Root Cause' ";
            sql = sql + ",[Releated Hardware] as Hardware ";
            sql = sql + ",[ปัญหาเกิดที่] as ปัญหาที่เกิด ";
            sql = sql + ",(select (case when[Root Cause] in ('OFC','BGP / Link Flap','Change Config','Network Improvement','Hardware Failure','Lost / Unstable Connection','Hardware Error (Hard - Reset)','Hardware Error (Soft - Reset)') ";
            sql = sql + "and (select (Datediff(mi,a.t1,a.t2)) from ";
            sql = sql + "(select CAST(ISNULL (CONVERT(DATE,[Link Down],102),CONVERT(DATE,[Create Date],102)) AS DATETIME) + CAST(ISNULL (CONVERT(TIME,[Link Down (Time)],108),CONVERT(TIME,[Create Date],108)) AS DATETIME) t1 ";
            sql = sql + ",CAST(ISNULL (CONVERT(DATE,[Link Up],102),CONVERT(DATE,[Closed Date],102)) AS DATETIME) + CAST(ISNULL (CONVERT(TIME,[Link Up (Time)],108),CONVERT(TIME,[Closed Date],108)) AS DATETIME) t2 ";
            sql = sql + ") a) > [min] ";
            sql = sql + "and [เอกสารใบเลื่อน] in ('','ไม่มีเอกสารใบเลื่อน') then 'Breach' ";
            sql = sql + "else 'Meet' end)) as 'Breach/Meet' ";
            sql = sql + ",[เอกสารใบเลื่อน]+' '+[OFC ขาดเนื่องจากสาเหตุ] as 'ข้อยกเว้น' ";
            sql = sql + ",[วิเคราะห์ Customer] as 'วิเคราะห์ Customer' ";
            sql = sql + "from [dbGIN].[dbo].[Ostickets2] LEFT JOIN[dbGIN].[dbo].[Sla] ON[dbGIN].[dbo].[Sla].slamin = [dbGIN].[dbo].[Ostickets2].SLA ";
            sql = sql + "where '" + txtDate1 + "' <= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) and '" + txtDate2 + "' >= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) ";
            sql = sql + "order by [Create Date] ";

            if (cn.State == ConnectionState.Open)
                cn.Close();
            cn.ConnectionString = strConn;
            cn.Open();

            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(sql, cn);
            da.Fill(dt);
            GridView1.DataSource = dt;
            GridView1.DataBind();
            cn.Close();

            String txtDate3 = TextBox3.Text;
            String txtDate4 = TextBox4.Text;

            sqln = "select	COALESCE(NULLIF(RIGHT([Subject],LEN([Subject])-charindex(')',[Subject])) , ''),SUBSTRING([Subject],1,case when charindex('(',[Subject]) = 0 then LEN([Subject]) else charindex('(',[Subject])-1 END) ) as หน่วยงานผู้ใช้ ";
            sqln = sqln + ",[Case ID]  as 'Case ID' ";
            sqln = sqln + ",[Created By] as ผู้แจ้งเหตุ ";
            sqln = sqln + ",ISNULL([เจ้าหน้าที่ปิดเคส Netka],[เจ้าหน้าที่ปิดเคส Netka (Other)]) as ผู้รับแจ้งเหตุ ";
            sqln = sqln + ",ISNULL (CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103))  as วันที่แจ้งเหตุ  ";
            sqln = sqln + ",ISNULL (CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) as เวลาแจ้งเหตุ ";
            sqln = sqln + ",[เหตุขัดข้อง] as เหตุขัดข้อง ";
            sqln = sqln + ",SUBSTRING([สรุปสาเหตุ # วิธีแก้ปัญหา],1,case when charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา]) = 0 then LEN([สรุปสาเหตุ # วิธีแก้ปัญหา]) else charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])-1 END) as สาเหตุ ";
            sqln = sqln + ",RIGHT([สรุปสาเหตุ # วิธีแก้ปัญหา],LEN([สรุปสาเหตุ # วิธีแก้ปัญหา])-charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])) as การแก้ไข ";
            sqln = sqln + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],103),CONVERT(VARCHAR(10),[Closed Date],103)) as วันที่แก้ไข ";
            sqln = sqln + ",ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) as เวลาแก้ไข ";
            sqln = sqln + ",(select	CONVERT(VARCHAR(10),Datediff(ss,a.t1,a.t2)/(60*60*24)) + ' Days ' ";
            sqln = sqln + "+CONVERT(VARCHAR(5),DateAdd(SS,Datediff(ss,a.t1, a.t2)%(60*60*24),0),114) as ระยะเวลาการแก้ไข ";
            sqln = sqln + "from";
            sqln = sqln + "(select ISNULL (CONVERT(VARCHAR(10),[Link Down],102),CONVERT(VARCHAR(10),[Create Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) t1 ";
            sqln = sqln + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],102),CONVERT(VARCHAR(10),[Closed Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) t2 ";
            sqln = sqln + ") a) as 'ระยะเวลาการแก้ไข' ";
            sqln = sqln + ",[dbGIN].[dbo].[Netka].[Priority] as ความรุนแรงของเหตุการ ";
            sqln = sqln + ",CONVERT(VARCHAR(3),DATENAME(dw,ISNULL ([Link Down],[Create Date]))) +'-'+CONVERT(VARCHAR(3),DATENAME(dw,ISNULL ([Link Up],[Closed Date]))) as 'หมายเหตุ' ";
            sqln = sqln + ",[dbGIN].[dbo].[Netka].[Root Cause] as 'Root Cause' ";
            sqln = sqln + ",[Case Category] as 'Case Category' ";
            sqln = sqln + ",[Case Sub Category] as 'Case Sub Category' ";
            sqln = sqln + "from [dbGIN].[dbo].[Ostickets2]  LEFT JOIN [dbGIN].[dbo].[Sla] ON [dbGIN].[dbo].[Sla].slamin = [dbGIN].[dbo].[Ostickets2].SLA ";
            sqln = sqln + "INNER JOIN [dbGIN].[dbo].[Netka] ON [dbGIN].[dbo].[Netka].[Case ID] = [dbGIN].[dbo].[Ostickets2].[DGA Ticket ID]	";
            sqln = sqln + "where '" + txtDate3 + "' <= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) and '" + txtDate4 + "' >= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) ";
            sqln = sqln + "order by[Create Date] ";

            if (cnn.State == ConnectionState.Open)
                cnn.Close();
            cnn.ConnectionString = strConn;
            cnn.Open();

            DataTable dtn = new DataTable();
            SqlDataAdapter dan = new SqlDataAdapter(sqln, cnn);
            dan.Fill(dtn);
            GridView2.DataSource = dtn;
            GridView2.DataBind();
            cnn.Close();

        }
        public override void VerifyRenderingInServerForm(Control control)
        {
            //base.VerifyRenderingInServerForm(control);
        }
        protected void Button1_Click(object sender, EventArgs e)
        {

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment;filename={0}", "SLA_report.xls"));
            Response.ContentType = "application/ms-excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            GridView1.AllowPaging = false;
            bindgrid();
            GridView1.HeaderRow.Style.Add("background-color", "#ffff");
            for (int i = 0; i < GridView1.HeaderRow.Cells.Count; i++)
            {
                GridView1.HeaderRow.Cells[i].Style.Add("background-color", "#FFA500");
            }
            GridView1.RenderControl(hw);
            Response.Write(sw.ToString());
            Response.End();
        }
        protected void Button2_Click(object sender, EventArgs e)
        {

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment;filename={0}", "Netka_report.xls"));
            Response.ContentType = "application/ms-excel";
            StringWriter swn = new StringWriter();
            HtmlTextWriter hwn = new HtmlTextWriter(swn);
            GridView2.AllowPaging = false;
            bindgrid();
            GridView2.HeaderRow.Style.Add("background-color", "#ffff");
            for (int i = 0; i < GridView2.HeaderRow.Cells.Count; i++)
            {
                GridView2.HeaderRow.Cells[i].Style.Add("background-color", "#B0E0E6");
            }
            GridView2.RenderControl(hwn);
            Response.Write(swn.ToString());
            Response.End();
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            String Ticket_Number;
            DateTime Create_Date;
            DateTime Lastresponse;
            DateTime? Closed_Date = null;
            String Subject;
            String From;
            String From_Email;
            String Priority;
            float priority_id;
            String Department;
            String Help_Topic;
            String Source;
            String Current_Status;
            String SLA_Period;
            DateTime Last_Updated;
            DateTime? Due_Date = null;
            float Overdue;
            float Answered;
            String Assigned_To;
            String Agent_Assigned;
            String Team_Assigned;
            String Ticket_Source_by;
            String Ticket_Source_from;
            String Circuit_ID;
            String Project_Code;
            String AIT_Ticket_ID;
            String DGA_Ticket_ID;
            String SCOM_Ticket_ID;
            String จังหวัด;
            String SLA;
            String เหตุขัดข้อง;
            String เหตุขัดข้อง_อื่นๆ;
            DateTime? Link_Down = null;
            DateTime? Link_Down_Time = null;
            String Close_Case_by;
            String Forward_Case_To;
            String ช่างที่ดำเนินการแก้ไข;
            String ชื่อ_นามสกุล_ช่าง;
            String เบอร์ติดต่อช่าง;
            String Appointed_time;
            String สรุปสาเหตุ_วิธีแก้ปัญหา;
            String OFC_ขาดเนื่องจากสาเหตุ;
            String วิเคราะห์_Customer;
            DateTime? Link_Up = null;
            DateTime? Link_Up_Time = null;
            String Root_Cause;
            String Root_Cause_Other;
            String ปัญหาเกิดที่;
            String Releated_Hardware;
            String Releated_Hardware_Other;
            String SN_ตัวเสีย;
            String SN_ตัวใหม่;
            String เอกสารใบเลื่อน;
            String ใบเลื่อนโดย;
            String เจ้าหน้าที่ปิดเคส_Netka;
            String เจ้าหน้าที่ปิดเคส_Netka_Other;

            string path = Path.GetFileName(FileUpload1.FileName);
            path = path.Replace(" ", "");
            FileUpload1.SaveAs(Server.MapPath("~/ExcelFile/") + path);
            String ExcelPath = Server.MapPath("~/ExcelFile/") + path;
            OleDbConnection mycon = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties=Excel 8.0; Persist Security Info = False");
            mycon.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", mycon);
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                // Response.Write("<br/>"+dr[0].ToString());
                //Roll_No = Convert.ToInt32(dr[0].ToString());
                Ticket_Number = dr[0].ToString();
                Create_Date = Convert.ToDateTime(dr[1].ToString());
                Lastresponse = Convert.ToDateTime(dr[2].ToString());
                if (string.IsNullOrEmpty(dr[3].ToString()))
                {
                    Closed_Date = null;
                }
                else
                {
                    Closed_Date = Convert.ToDateTime(dr[3].ToString());
                }
                Subject = dr[4].ToString();
                From = dr[5].ToString();
                From_Email = dr[6].ToString();
                Priority = dr[7].ToString();
                priority_id = Convert.ToInt32(dr[8].ToString());
                Department = dr[9].ToString();
                Help_Topic = dr[10].ToString();
                Source = dr[11].ToString();
                Current_Status = dr[12].ToString();
                SLA_Period = dr[13].ToString();
                Last_Updated = Convert.ToDateTime(dr[14].ToString());
                if (string.IsNullOrEmpty(dr[15].ToString()))
                {
                    Due_Date = null;
                }
                else
                {
                    Due_Date = Convert.ToDateTime(dr[15].ToString());
                }
                Overdue = Convert.ToInt32(dr[16].ToString());
                Answered = Convert.ToInt32(dr[17].ToString());
                Assigned_To = dr[18].ToString();
                Agent_Assigned = dr[19].ToString();
                Team_Assigned = dr[20].ToString();
                Ticket_Source_by = dr[21].ToString();
                Ticket_Source_from = dr[22].ToString();
                Circuit_ID = dr[23].ToString();
                Project_Code = dr[24].ToString();
                AIT_Ticket_ID = dr[25].ToString();
                DGA_Ticket_ID = dr[26].ToString();
                SCOM_Ticket_ID = dr[27].ToString();
                จังหวัด = dr[28].ToString();
                SLA = dr[29].ToString();
                เหตุขัดข้อง = dr[30].ToString();
                เหตุขัดข้อง_อื่นๆ = dr[31].ToString();
                if (string.IsNullOrEmpty(dr[32].ToString()))
                {
                    Link_Down = null;
                }
                else
                {

                    Link_Down = Convert.ToDateTime(dr[32].ToString());
                }

                if (string.IsNullOrEmpty(dr[33].ToString()))
                {
                    Link_Down_Time = null;
                }
                else
                {

                    Link_Down_Time = Convert.ToDateTime(dr[33].ToString());
                }

                //Link_Down = Convert.ToDateTime(dr[32].ToString());
                //Link_Down_Time = Convert.ToDateTime(dr[33].ToString());
                Close_Case_by = dr[34].ToString();
                Forward_Case_To = dr[35].ToString();
                ช่างที่ดำเนินการแก้ไข = dr[36].ToString();
                ชื่อ_นามสกุล_ช่าง = dr[37].ToString();
                เบอร์ติดต่อช่าง = dr[38].ToString();
                Appointed_time = dr[39].ToString();
                สรุปสาเหตุ_วิธีแก้ปัญหา = dr[40].ToString();
                OFC_ขาดเนื่องจากสาเหตุ = dr[41].ToString();
                วิเคราะห์_Customer = dr[42].ToString();
                if (string.IsNullOrEmpty(dr[43].ToString()))
                {
                    Link_Up = null;
                }
                else
                {

                    Link_Up = Convert.ToDateTime(dr[43].ToString());
                }
                if (string.IsNullOrEmpty(dr[44].ToString()))
                {
                    Link_Up_Time = null;
                }
                else
                {

                    Link_Up_Time = Convert.ToDateTime(dr[44].ToString());
                }

                Root_Cause = dr[45].ToString();
                Root_Cause_Other = dr[46].ToString();
                ปัญหาเกิดที่ = dr[47].ToString();
                Releated_Hardware = dr[48].ToString();
                Releated_Hardware_Other = dr[49].ToString();
                SN_ตัวเสีย = dr[50].ToString();
                SN_ตัวใหม่ = dr[51].ToString();
                เอกสารใบเลื่อน = dr[52].ToString();
                ใบเลื่อนโดย = dr[53].ToString();
                เจ้าหน้าที่ปิดเคส_Netka = dr[54].ToString();
                เจ้าหน้าที่ปิดเคส_Netka_Other = dr[55].ToString();
                savedata(Ticket_Number, Create_Date, Lastresponse, Closed_Date, Subject, From, From_Email, Priority, priority_id, Department, Help_Topic, Source, Current_Status, SLA_Period, Last_Updated, Due_Date, Overdue
                    , Answered, Assigned_To, Agent_Assigned, Team_Assigned, Ticket_Source_by, Ticket_Source_from, Circuit_ID, Project_Code, AIT_Ticket_ID, DGA_Ticket_ID, SCOM_Ticket_ID, จังหวัด, SLA, เหตุขัดข้อง, เหตุขัดข้อง_อื่นๆ
                    , Link_Down, Link_Down_Time, Close_Case_by, Forward_Case_To, ช่างที่ดำเนินการแก้ไข, ชื่อ_นามสกุล_ช่าง, เบอร์ติดต่อช่าง, Appointed_time, สรุปสาเหตุ_วิธีแก้ปัญหา, OFC_ขาดเนื่องจากสาเหตุ, วิเคราะห์_Customer, Link_Up, Link_Up_Time, Root_Cause
                    , Root_Cause_Other, ปัญหาเกิดที่, Releated_Hardware, Releated_Hardware_Other, SN_ตัวเสีย, SN_ตัวใหม่, เอกสารใบเลื่อน, ใบเลื่อนโดย, เจ้าหน้าที่ปิดเคส_Netka, เจ้าหน้าที่ปิดเคส_Netka_Other);


            }
            Label1.Text = "Data Has Been Saved Successfully";

        }
        private void savedata(String Ticket_Number1, DateTime Create_Date1, DateTime Lastresponse1, DateTime? Closed_Date1, String Subject1, String From1, String From_Email1, String Priority1,
            float priority_id1, String Department1, String Help_Topic1, String Source1, String Current_Status1, String SLA_Period1, DateTime Last_Updated1, DateTime? Due_Date1, float Overdue1
            , float Answered1, String Assigned_To1, String Agent_Assigned1, String Team_Assigned1, String Ticket_Source_by1, String Ticket_Source_from1, String Circuit_ID1, String Project_Code1
            , String AIT_Ticket_ID1, String DGA_Ticket_ID1, String SCOM_Ticket_ID1, String จังหวัด1, String SLA1, String เหตุขัดข้อง1, String เหตุขัดข้อง_อื่นๆ1, DateTime? Link_Down1, DateTime? Link_Down_Time1
            , String Close_Case_by1, String Forward_Case_To1, String ช่างที่ดำเนินการแก้ไข1, String ชื่อ_นามสกุล_ช่าง1, String เบอร์ติดต่อช่าง1, String Appointed_time1, String สรุปสาเหตุ_วิธีแก้ปัญหา1, String OFC_ขาดเนื่องจากสาเหตุ1
            , String วิเคราะห์_Customer1, DateTime? Link_Up1, DateTime? Link_Up_Time1, String Root_Cause1, String Root_Cause_Other1, String ปัญหาเกิดที่1, String Releated_Hardware1, String Releated_Hardware_Other1
            , String SN_ตัวเสีย1, String SN_ตัวใหม่1, String เอกสารใบเลื่อน1, String ใบเลื่อนโดย1, String เจ้าหน้าที่ปิดเคส_Netka1, String เจ้าหน้าที่ปิดเคส_Netka_Other1)
        {
            String query = "insert into Ostickets2([Ticket Number],[Create Date],[Lastresponse],[Subject],[From],[From Email],[Priority],[priority_id],[Department],[Help Topic],[Source],[Current Status],[SLA Period]" +
                ",[Last Updated],[Overdue],[Answered],[Assigned To],[Agent Assigned],[Team Assigned],[Ticket Source by],[Ticket Source from],[Circuit ID],[Project Code],[AIT Ticket ID],[DGA Ticket ID],[SCOM Ticket ID]" +
                ",[จังหวัด],[SLA],[เหตุขัดข้อง],[เหตุขัดข้อง (อื่นๆ)],[Close Case by],[Forward Case To],[ช่างที่ดำเนินการแก้ไข],[ชื่อ - นามสกุล ช่าง],[เบอร์ติดต่อช่าง],[Appointed time],[สรุปสาเหตุ # วิธีแก้ปัญหา],[OFC ขาดเนื่องจากสาเหตุ]" +
                ",[วิเคราะห์ Customer],[Root Cause],[Root Cause (Other)],[ปัญหาเกิดที่],[Releated Hardware],[Releated Hardware (Other)],[S/N (ตัวเสีย)],[S/N (ตัวใหม่)],[เอกสารใบเลื่อน],[ใบเลื่อนโดย],[เจ้าหน้าที่ปิดเคส Netka],[เจ้าหน้าที่ปิดเคส Netka (Other)]) " +
                "values('" + Ticket_Number1 + "',convert(datetime, '" + Create_Date1 + "', 103),convert(datetime, '" + Lastresponse1 + "', 103),'" + Subject1 + "','" + From1 + "','" + From_Email1 + "','" + Priority1 + "','" + priority_id1 + "','" + Department1 +
                "','" + Help_Topic1 + "','" + Source1 + "','" + Current_Status1 + "','" + SLA_Period1 + "',convert(datetime, '" + Last_Updated1 + "', 103),'" + Overdue1 + "','" + Answered1 + "','" + Assigned_To1 + "','" + Agent_Assigned1 + "','"
                + Team_Assigned1 + "','" + Ticket_Source_by1 + "','" + Ticket_Source_from1 + "','" + Circuit_ID1 + "','" + Project_Code1 + "','" + AIT_Ticket_ID1 + "','" + DGA_Ticket_ID1 + "','" + SCOM_Ticket_ID1 + "','" + จังหวัด1 + "','"
                + SLA1 + "','" + เหตุขัดข้อง1 + "','" + เหตุขัดข้อง_อื่นๆ1 + "','" + Close_Case_by1 + "','" + Forward_Case_To1 + "','" + ช่างที่ดำเนินการแก้ไข1 + "','" + ชื่อ_นามสกุล_ช่าง1 + "','" + เบอร์ติดต่อช่าง1 + "','"
                + Appointed_time1 + "','" + สรุปสาเหตุ_วิธีแก้ปัญหา1 + "','" + OFC_ขาดเนื่องจากสาเหตุ1 + "','" + วิเคราะห์_Customer1 + "','" + Root_Cause1 + "','" + Root_Cause_Other1 + "','" + ปัญหาเกิดที่1 + "','"
                + Releated_Hardware1 + "','" + Releated_Hardware_Other1 + "','" + SN_ตัวเสีย1 + "','" + SN_ตัวใหม่1 + "','" + เอกสารใบเลื่อน1 + "','" + ใบเลื่อนโดย1 + "','" + เจ้าหน้าที่ปิดเคส_Netka1 + "','" + เจ้าหน้าที่ปิดเคส_Netka_Other1 + "')";
            //String mycon = "Data Source=localhost\sqlexpress;Initial Catalog=dbGIN;Integrated Security=True";
            String mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
            SqlConnection con = new SqlConnection(mycon);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = query;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();
            con.Close();

            if (Closed_Date1 != null)
            {
                String queryclosedate = "update Ostickets2 set [Closed Date] = convert(datetime,'" + Closed_Date1 + "',103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = queryclosedate;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }

            if (Due_Date1 != null)
            {
                String queryduedate = "update Ostickets2 set [Due Date] = convert(datetime,'" + Due_Date1 + "',103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";

                //String mycon = "Data Source=localhost\sqlexpress;Initial Catalog=dbGIN;Integrated Security=True";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = queryduedate;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            if (Link_Down1 != null)
            {
                String querylinkdown = "update Ostickets2 set [Link Down] = convert(datetime, '" + Link_Down1 + "', 103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";

                //String mycon = "Data Source=localhost\sqlexpress;Initial Catalog=dbGIN;Integrated Security=True";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = querylinkdown;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            if (Link_Up1 != null)
            {
                String querylinkup = "update Ostickets2 set [Link Up] = convert(datetime, '" + Link_Up1 + "', 103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = querylinkup;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            if (Link_Down_Time1 != null)
            {
                String querylinkdowntime = "update Ostickets2 set [Link Down (Time)] = convert(datetime, '" + Link_Down_Time1 + "', 103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = querylinkdowntime;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }
            if (Link_Up_Time1 != null)
            {
                String querylinkuptime = "update Ostickets2 set [Link Up (Time)] = convert(datetime, '" + Link_Up_Time1 + "', 103) where  [Ticket Number] = ('" + Ticket_Number1 + "') ";
                mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
                con = new SqlConnection(mycon);
                con.Open();
                cmd = new SqlCommand();
                cmd.CommandText = querylinkuptime;
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
            }


        }

    }
}
