//Main Window Logic
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Controls;
using System.IO;

namespace AbandonedCalls2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //SQL Connection to Mitel
        SqlConnection conn = new SqlConnection("Data Source=server;" +
                "Persist Security Info=True;User ID=username;Password=password");

        //DataTable Creation
        DataTable CallsTable = new DataTable();

        //List to pull from
        List<Calls> FirstList = new List<Calls>();

        //List to save to
        List<MailObjects> ThingsToMail = new List<MailObjects>();

        //Marker Numeral
        int mrk = 0;

        //Invoking Email Application & Mail Message
        Outlook.Application oApp = new Outlook.Application();
        Email cMsg = new Email();


        public MainWindow()
        {
            InitializeComponent();

            conn.Open();
            SqlDataAdapter script = new SqlDataAdapter
                ("declare @end_date datetime " +
                "declare @start_date datetime " +
                "declare @today1 datetime " +
                "set @today1 = DATEADD(Day, 0, DATEDIFF(Day, 0, GetDate())) " +
                "set @end_date = @today1+1 " +
                "set @start_date = @today1 " +
                "SELECT " +
                "midnightstartdate, fullani , timetoabandon, name " +
                "FROM [CCMData].[dbo].[tblData_QueueAbandonByANI] " +
                "join  [CCMData].[dbo].[tblConfig_Queue] on " +
                "[CCMData].[dbo].[tblData_QueueAbandonByANI].fkqueue " +
                "=  [CCMData].[dbo].[tblConfig_Queue].pkey " +
                "where FKQueue in ('A6C60267-D717-430E-9B34-010BC621C41A','17089379-371A-436F-A787-5119662A9431','A812DC10-07D9-4158-A4BD-FE254204D77B','8D9FE43E-4CC0-461E-9C65-9BA38B23DB59')" +
                "and  MidnightStartDate between @start_date and @end_date " +
                "and fullani not in ('T1', '5146695500') " +
                "and fullani not in " +
                "(  select ani from [CCMData].[dbo].[tblData_InboundTrace] " +
                "join  [CCMData].[dbo].[tblConfig_Queue] on [CCMData].[dbo].[tblData_InboundTrace].fkqueue " +
                "=  [CCMData].[dbo].[tblConfig_Queue].pkey " +
                "where FKQueue in ('A6C60267-D717-430E-9B34-010BC621C41A','17089379-371A-436F-A787-5119662A9431','A812DC10-07D9-4158-A4BD-FE254204D77B','8D9FE43E-4CC0-461E-9C65-9BA38B23DB59') " +
                "and  MidnightStartDate between @start_date and @end_date " +
                "and duration >1) "
                , conn);

            script.Fill(CallsTable);

            //temp variable for each new item in calls object.
            ListViewItem tempItem=new ListViewItem();

            //Displays Abandoned Calls Information
            for (int i = 0; i < CallsTable.Rows.Count - 1; i++)
            {
                Calls c = new Calls(CallsTable.Rows[i][0].ToString(), CallsTable.Rows[i][3].ToString(), CallsTable.Rows[i][2].ToString(), CallsTable.Rows[i][1].ToString(), " ");
                //tempItem.Content = c;
                FirstList.Add(c);

            }
            FirstList = FirstList.OrderBy(Calls => Calls.WhenCalled).ToList();
            Numberly();
            Orderly();

            //Buttons
            GoButton.Click += GoButton_Click;
            ExitButton.Click += ExitButton_Click;
            SendEmail.Click += SendEmail_Click;

            //Error Log Writer
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);
        }

        private void SendEmail_Click(object sender, RoutedEventArgs e)
        {
            //Constructs the mail message
            MailBuilder();
            //Sets and Sends the Mail message (and clears it)
            Mailman();

            SendEmail.IsEnabled = false;
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            //Script Cleanup and Quit
            Janitor();
            Close();
        }

        private void GoButton_Click(object sender, RoutedEventArgs e)
        {
            //Main Calling Function
            int thisCall = callsListView.SelectedIndex;

            if (optsBox.Text != "<none>")
                FirstList[thisCall].Status = optsBox.Text;
            else
                FirstList[thisCall].Status = " ";

            MailObjects mO = new MailObjects();
            mO.timeOfCall = FirstList[thisCall].WhenCalled;
            mO.numOfCall = FirstList[thisCall].ContactNum;
            mO.callResult = optsBox.Text;

            if (FirstList[thisCall].Status != " " && !ThingsToMail.Contains(mO))
            {
                ThingsToMail.Add(mO);
                mrk++;
            }
            else if (FirstList[thisCall].Status == " ")
            {
                try
                {
                    //int fr = ThingsToMail.IndexOf(MailObjects => MailObjects.numOfCall == mO.numOfCall);
                    //ObjectsList.Where((v, i) => v.StringInObject == NamedObject.StringInObject ? i : -1).FirstOrDefault(v => v >= 0);
                    ThingsToMail.RemoveAll(MailObjects => MailObjects.numOfCall == mO.numOfCall);
                }
                catch (ArgumentNullException)
                {
                }

                mrk--;
            }

            mO = null;

            if (mrk >= FirstList.Count)
            {
                SendEmail.IsEnabled = true;
                HiddenInfo2.Content = "List Completed";
            }
            else
                SendEmail.IsEnabled = false;

            optsBox.SelectedIndex = -1;
            Orderly();
            callsListView.SelectedIndex = 0;

        }

        private void Janitor()
        {
        //Cleanup to prevent Memory Leaking
            oApp = null;
            cMsg = null;
            FirstList = null;
            CallsTable = null;
            ThingsToMail = null;
        }

        private void Mailman()
        {
        //Sends the Mail message
            Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            oMsg.Recipients.Add(cMsg.Recipient);
            oMsg.Subject = cMsg.Subject;
            oMsg.HTMLBody += cMsg.Body;
            oMsg.HTMLBody += cMsg.htmlTable;
            oMsg.HTMLBody += "<br>";
            oMsg.HTMLBody += "<p><small>This message was automatically Generated by ForgetMeNot2</small></p>";

            oMsg.Send();

            //Cleanup
            oMsg = null;
        }

        public void MailBuilder()
        {
            //Builds the Mail Message
            cMsg.Recipient = "CDN-helpdesk@moneymart.ca";
            cMsg.Subject = "Abandoned Calls Follow Up for " + DateTime.Now;
            cMsg.Body += "<p>Hello, what follows is today's Abandoned Call Follow-Ups.</p>";

            cMsg.htmlTable += "<table style=\"width:85%\"><tr style=\"color:#4B67D6\"><th>Time Called</th><th>Number</th><th>Contacted</th></tr>";
            for (int e = 0; e < mrk; e++)
            {
                cMsg.htmlTable += "<tr><td>" + ThingsToMail[e].timeOfCall + "</td><td>" + ThingsToMail[e].numOfCall + "</td><td>" + ThingsToMail[e].callResult + "</td></tr>";
            }
            cMsg.htmlTable += "</table>";

        }

        public void Numberly()
        {
            //Cleans up Phone Numbers for View
            for (int i = 0; i < FirstList.Count; i++)
            {
                string Numerary = FirstList[i].ContactNum;

                if (FirstList[i].ContactNum.Length > 6 && FirstList[i].ContactNum.Length < 11)
                {
                    Numerary = Numerary.Insert(3, "-");
                    Numerary = Numerary.Insert(7, "-");
                    FirstList[i].ContactNum = Numerary;
                }
                else if (FirstList[i].ContactNum.Length > 10 && FirstList[i].ContactNum.Length < 14)
                {
                    Numerary = Numerary.Insert(1, "-");
                    Numerary = Numerary.Insert(5, "-");
                    Numerary = Numerary.Insert(9, "-");
                    FirstList[i].ContactNum = Numerary;
                }
            }
        }

        public void Orderly()
        {
            //Reorders the list as calls are completed.
            callsListView.Items.Clear();

            for (int f = 0; f < FirstList.Count; f++)
            {
                var tempItem2 = FirstList[f];

                if (FirstList[f].Status != " ")
                {
                    FirstList.RemoveAt(f);
                    //FirstList.Insert(FirstList.Count, tempItem2);
                    FirstList.Add(tempItem2);
                }

                callsListView.Items.Add(FirstList[f]);
            }

            int cnt = FirstList.Count;

            HiddenInfo.Content = "Calls Total: " + mrk + "/" + cnt;
        }

        private void TheWatcher()
        {
        //Watcher for Button Enabler
            if (optsBox.SelectedIndex != -1 && callsListView.SelectedIndex != -1)
                GoButton.IsEnabled = true;
            else
                GoButton.IsEnabled = false;
        }

        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
        //Simple Bug Reporting
            Exception e = (Exception)args.ExceptionObject;

            using (StreamWriter sw = new StreamWriter(@"L:\Nic\ForgetMeNot\Error_Log.txt", true))
            {
                sw.WriteLine(e);
                sw.WriteLine(DateTime.Now);
                sw.WriteLine("-----------------------------------------------------------------------------");
            }

        }

        private void optsBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TheWatcher();
        }
    }
    
    public class MailObjects
    {
        public string timeOfCall { get; set; }
        public string numOfCall { get; set; }
        public string callResult { get; set; }
    }
}
