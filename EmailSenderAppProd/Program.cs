using System;
using System.IO;
using System.Data;
using Quartz;
using Quartz.Impl;
using System.Data.SqlClient;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Text;
using HandlebarsDotNet;
using System.Linq;

namespace EmailSenderApp
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            IScheduler scheduler = await InitializeScheduler();

            await ScheduleJob(scheduler);

            Console.ReadLine();
            await ShutdownScheduler(scheduler);

            ApplicationShutdown.CloseApplication();
        }

        static async System.Threading.Tasks.Task<IScheduler> InitializeScheduler()
        {
            StdSchedulerFactory factory = new StdSchedulerFactory();
            IScheduler scheduler = await factory.GetScheduler();
            await scheduler.Start();
            return scheduler;
        }

        static async System.Threading.Tasks.Task ScheduleJob(IScheduler scheduler)
        {
            IJobDetail job = CreateJob<SendMail>("myJob", "group1");
            DateTime desiredDateTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 3, 10, 0, 0);
            TimeSpan initialDelay = desiredDateTime - DateTime.Now;
            ITrigger trigger = CreateTrigger("myTrigger", "group1", DateTimeOffset.Now.Add(initialDelay));

            await scheduler.ScheduleJob(job, trigger);
        }

        static async System.Threading.Tasks.Task ShutdownScheduler(IScheduler scheduler)
        {
            await scheduler.Shutdown();
        }

        static IJobDetail CreateJob<T>(string name, string group) where T : IJob
        {
            return JobBuilder.Create<T>().WithIdentity(name, group).Build();
        }

        static ITrigger CreateTrigger(string name, string group, DateTimeOffset startTime)
        {
            return TriggerBuilder.Create().WithIdentity(name, group).StartAt(startTime).Build();
        }
    }

    public class SendMail : IJob
    {
        private readonly ExchangeService _service;

        public SendMail()
        {
            _service = InitializeExchangeService();
        }

        public async System.Threading.Tasks.Task Execute(IJobExecutionContext context)
        {
            await Console.Out.WriteLineAsync("Email Scheduler Starting at: " + DateTime.Now.ToString());

            string connectionString = ConfigurationManager.ConnectionStrings["connStringDb"].ConnectionString;
            ExchangeService service = InitializeExchangeService();

            string logFilePath = ConfigurationManager.AppSettings["LogFilePath"];
            EmailLogger emailLogger = new EmailLogger(logFilePath);

            try
            {
                ApplicationShutdown.IncrementSendmailProcess();

                DataTable dt = ExecuteStoredProcedure(connectionString);
                var groupedData = GroupDataByBranches(dt);

                foreach (var group in groupedData)
                {
                    string branches = group.Key;
                    string dueDates = string.Join(", ", group.Select(row => row.Field<string>("DUEDATE")));
                    List<List<dynamic>> groupValuesList = GetValuesFromDataTable(group.ToList());

                    string templatePath = "D:/emailTemplate/EmailTemplate.html";
                    string template = File.ReadAllText(templatePath);
                    string htmlBody = CreateTableContent(template, groupValuesList.SelectMany(x => x).ToList());

                    EmailMessage message = CreateEmailMessage(htmlBody, branches, dueDates);
                    AddRecipients(message, group);
                    AddAttachments(message);

                    message.SendAndSaveCopy();

                    emailLogger.LogEmail(branches, group.Select(row => row.Field<string>("DUEDATE")).ToArray(), group.Select(row => row.Field<string>("TOPICS")).ToList());

                    Console.WriteLine("Email sent successfully at: " + DateTime.Now.ToString());
                }
                ApplicationShutdown.DecrementSendmailProcess();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred: " + ex.Message);
                ApplicationShutdown.DecrementSendmailProcess();
            }
        }

        private ExchangeService InitializeExchangeService()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            string username = ConfigurationManager.AppSettings["Username"];
            string password = ConfigurationManager.AppSettings["Password"];
            string email = ConfigurationManager.AppSettings["Email"];

            service.Credentials = new WebCredentials(username, password);
            service.AutodiscoverUrl(email);

            return service;
        }

        static DataTable ExecuteStoredProcedure(string connectionString)
        {
            string storedProc = "SP_GETDATA_TABLE";
            using (SqlConnection sqlConn = new SqlConnection(connectionString))
            {
                SqlCommand sqlGetData = new SqlCommand(storedProc, sqlConn);
                sqlGetData.CommandType = CommandType.StoredProcedure;

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(sqlGetData);

                sqlConn.Open();
                da.Fill(dt);
                return dt;
            }
        }

        static IEnumerable<IGrouping<string, DataRow>> GroupDataByBranches(DataTable dt)
        {
            if (dt != null)
            {
                return dt.AsEnumerable().GroupBy(row => row.Field<string>("BRANCHES"));
            }
            return Enumerable.Empty<IGrouping<string, DataRow>>();
        }

        static List<List<dynamic>> GetValuesFromDataTable(List<DataRow> rows)
        {
            List<List<dynamic>> valuesList = new List<List<dynamic>>();
            foreach (var row in rows)
            {
                List<dynamic> values = new List<dynamic>
                {
                    new
                    {
                        topic = row.Field<string>("TOPICS"),
                        branches = row.Field<string>("BRANCHES"),
                        branchLeader = row.Field<string>("BRANCH_LEADER_EMAIL"),
                        branchCoLeader = row.Field<string>("BRANCH_COLEADER_EMAIL"),
                        duedate = row.Field<string>("DUE_DATES")
                    }
                };
                valuesList.Add(values);
            }
            return valuesList;
        }

        static string CreateTableContent(string template, List<dynamic> values)
        {
            StringBuilder tableBuilder = new StringBuilder();
            tableBuilder.Append("<table role=\"presentation\" style=\"width: 100%; border: 1px solid #ccc; border-spacing: 0; table-layout: auto;\">");
            tableBuilder.Append("<thead style=\"font-size: 1.2em; line-height: 1.5em; font-family: Arial, sans-serif; text-align: center;\">");
            tableBuilder.Append("<tr>");
            tableBuilder.Append("<th style=\"border: 1px solid; border-spacing: 0;\">Topics</th>");
            tableBuilder.Append("<th style=\"border: 1px solid; border-spacing: 0;\">BRANCHES</th>");
            tableBuilder.Append("<th style=\"border: 1px solid; border-spacing: 0;\">Due Date</th>");
            tableBuilder.Append("</tr>");
            tableBuilder.Append("</thead>");
            tableBuilder.Append("<tbody>");

            foreach (var item in values)
            {
                tableBuilder.Append("<tr>");
                tableBuilder.Append("<td style=\"border: 1px solid; border-spacing: 0; vertical-align:top;\">" + item.topic + "</td>");
                tableBuilder.Append("<td style=\"border: 1px solid; border-spacing: 0; vertical-align:top;\">" + item.branches + "</td>");
                tableBuilder.Append("<td style=\"border: 1px solid; border-spacing: 0; vertical-align:top;\">" + item.duedate + "</td>");
                tableBuilder.Append("</tr>");
            }

            tableBuilder.Append("</tbody>");
            tableBuilder.Append("</table>");

            return template.Replace("{{tableContent}}", tableBuilder.ToString());
        }

        private EmailMessage CreateEmailMessage(string htmlBody, string branches, string dueDates)
        {
            EmailMessage message = new EmailMessage(_service);

            StringBuilder bodyBuilder = new StringBuilder(htmlBody);

            bodyBuilder.Replace("{BRANCHES}", branches);
            bodyBuilder.Replace("{DUEDATE}", dueDates);

            message.Subject = "This is Email Reminder for delayed job of : " + branches;
            message.Body = new MessageBody(BodyType.HTML, bodyBuilder.ToString());

            return message;
        }

        static void AddRecipients(EmailMessage message, IGrouping<string, DataRow> group)
        {
            string BRANCHES = group.Key;

            HashSet<string> addedEmails = new HashSet<string>(); // Keep track of added email addresses

            var toBranchLeader = group.Select(row => new EmailAddress(row.Field<string>("BRANCH_LEADER_EMAIL"))).ToList();
            foreach (var email in toBranchLeader)
            {
                if (!addedEmails.Contains(email.Address))
                {
                    message.CcRecipients.Add(email);
                    addedEmails.Add(email.Address);
                }
            }

            var toBranchCoLeader = group.Select(row => new EmailAddress(row.Field<string>("BRANCH_COLEADER_EMAIL"))).ToList();
            foreach (var email in toBranchCoLeader)
            {
                if (!addedEmails.Contains(email.Address))
                {
                    message.CcRecipients.Add(email);
                    addedEmails.Add(email.Address);
                }
            }

            bool IsValidEmail(string email)
            {
                string emailPattern = @"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}$";
                return System.Text.RegularExpressions.Regex.IsMatch(email, emailPattern);
            }
        }

        static void AddAttachments(EmailMessage message)
        {
            string file = @"\emailTemplate\images\some-logo.png";
            FileAttachment attachment = message.Attachments.AddFileAttachment(file);
            attachment.ContentId = "some-logo";

            string file2 = @"\images\logo-b.png";
            FileAttachment attachment2 = message.Attachments.AddFileAttachment(file2);
            attachment2.ContentId = "logo-b";
        }
    }

    public class EmailLogger
    {
        private readonly string logFilePath;

        public EmailLogger(string logFilePath)
        {
            this.logFilePath = logFilePath;
        }

        public void LogEmail(string branches, string[] dueDates, List<string> topics)
        {
            string logMessage = $"Sending email at  " + DateTime.Now.ToString() + $" to BRANCHES: {branches} for Topics: {string.Join(", ", topics)} Due Dates: {string.Join(", ", dueDates)}";

            try
            {
                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine(logMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to log email: {ex.Message}");
            }
        }
    }
}
