using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using System.Xml.Xsl;

namespace MySchedule
{
    class Program
    {
        static void Main(string[] args)
        {
            string conString = ConfigurationManager.ConnectionStrings["constring"].ConnectionString;
            OleDbConnection con = new OleDbConnection(conString);
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataTable dt = new DataTable();

            string select = "select * from [Sheet1$]";
            cmd.Connection = con;
            cmd.CommandText = select;
            adapter.InsertCommand = cmd;
            adapter.SelectCommand = cmd;
            con.Open();
            adapter.Fill(dt);
            con.Close();

            DataTable scheduledt = new DataTable();
            scheduledt.TableName = "Table";

            DateTime today = DateTime.Now.Date;
            DateTime reference = new DateTime(2014, 09, 08);
            DayOfWeek dowToday = today.DayOfWeek;

            int diff = dowToday - DayOfWeek.Monday;

            if (diff < 0)
            {
                diff += 7;
            }

            DateTime monday;

            if (dowToday == DayOfWeek.Sunday)
            {
                monday = today.AddDays(1).Date;
            }
            else
            {
                monday = today.AddDays(-1 * diff).Date;
            }

            int scheduleNum = 1;

            if (Convert.ToInt32((monday - reference).Days) % 2 == 0)
            {
                scheduleNum = 1;
            }
            else
            {
                scheduleNum = 2;
            }

            int dtRowCount = 0;
            bool newLine = false;
            if (dowToday == DayOfWeek.Sunday)
            {
                dowToday++;
                scheduledt.Columns.Add("Day");
                scheduledt.Columns.Add("Time");
                scheduledt.Columns.Add("Schedule");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (Convert.ToInt32(dt.Rows[i][5].ToString()) == scheduleNum)
                    {
                        if (dtRowCount > 0)
                        {
                            if (scheduledt.Rows[dtRowCount - 1][0].ToString() != dt.Rows[i][4].ToString() && newLine == false)
                            {
                                scheduledt.Rows.Add();
                                scheduledt.Rows[dtRowCount][0] = "---";
                                scheduledt.Rows[dtRowCount][1] = "---";
                                scheduledt.Rows[dtRowCount][2] = "---";
                                i--;
                                newLine = true;
                            }
                            else
                            {
                                scheduledt.Rows.Add();
                                scheduledt.Rows[dtRowCount][0] = dt.Rows[i][4];
                                scheduledt.Rows[dtRowCount][1] = dt.Rows[i][2];
                                scheduledt.Rows[dtRowCount][2] = dt.Rows[i][0] + "     " + dt.Rows[i][1] + "     " + dt.Rows[i][3] + "     ";
                                newLine = false;
                            }
                        }
                        else if (dtRowCount == 0)
                        {
                            scheduledt.Rows.Add();
                            scheduledt.Rows[dtRowCount][0] = dt.Rows[i][4];
                            scheduledt.Rows[dtRowCount][1] = dt.Rows[i][2];
                            scheduledt.Rows[dtRowCount][2] = dt.Rows[i][0] + "     " + dt.Rows[i][1] + "     " + dt.Rows[i][3] + "     ";
                            newLine = false;
                        }
                        dtRowCount++;
                    }
                }
            }
            else
            {
                scheduledt.Columns.Add("Time");
                scheduledt.Columns.Add("Schedule");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][4].ToString() == dowToday.ToString() && Convert.ToInt32(dt.Rows[i][5].ToString()) == scheduleNum)
                    {
                        scheduledt.Rows.Add();
                        scheduledt.Rows[dtRowCount][0] = dt.Rows[i][2];
                        scheduledt.Rows[dtRowCount][1] = dt.Rows[i][0] + "     " + dt.Rows[i][1] + "     " + dt.Rows[i][3] + "     ";

                        dtRowCount++;
                    }
                }
            }

            string html = getHTML(scheduledt);

            Mail m1 = new Mail(html);

        }

        //construct HTML code
        static string getHTML(System.Data.DataTable dt)
        {
            string xsltFilePath = @"C:\Users\Kelly\Documents\schedule\MySchedule\MySchedule\bin\Debug\XSLTFile1.xslt";

            TextWriter txtwriter = new StringWriter();
            dt.WriteXml(txtwriter);
            XDocument xmlDoc = XDocument.Parse(txtwriter.ToString());

            XDocument htmlDoc = new XDocument();

            using (XmlWriter writer = htmlDoc.CreateWriter())
            {
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(xsltFilePath);
                xslt.Transform(xmlDoc.CreateReader(), writer);
            }

            return htmlDoc.Document.ToString();
        }

    }
}
