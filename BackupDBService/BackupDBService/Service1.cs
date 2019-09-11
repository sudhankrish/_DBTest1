using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using System.Configuration;
using Microsoft.Web.Administration;
using System.Net;


namespace BackupDBService
{
    public partial class BackupDBService : ServiceBase
    {
        private int interval;
        private Timer backupDBTimer = new Timer();

        public BackupDBService()
        {
            InitializeComponent();
        }

        protected override void OnContinue()
        {
            backupDBTimer.Start();
        }

        protected override void OnPause()
        {
            backupDBTimer.Stop();
        }

        protected override void OnStart(string[] args)
        {
            backupDBTimer.Enabled = true;
            DateTime currentTime = DateTime.Now;
            int intervalToElapse = 0;
            DateTime scheduleTime = Convert.ToDateTime(ConfigurationManager.AppSettings["TimeToRun"].ToString());

            if (currentTime <= scheduleTime)
                intervalToElapse = (int)scheduleTime.Subtract(currentTime).TotalSeconds;
            else
                intervalToElapse = (int)scheduleTime.AddDays(1).Subtract(currentTime).TotalSeconds;

            backupDBTimer = new System.Timers.Timer(intervalToElapse * 1000);
            backupDBTimer.AutoReset = true;
            backupDBTimer.Elapsed += new System.Timers.ElapsedEventHandler(Timeout);
            backupDBTimer.Start();
        }

        protected override void OnStop()
        {
            backupDBTimer.Stop();
        }

        private void Timeout(object sender, ElapsedEventArgs e)
        {
            try
            {
                // check if the time is correct if yes then do not set it
                DateTime signalTime = (e == null ? DateTime.UtcNow : e.SignalTime).ToUniversalTime();

                bool sucess = true;

                var server = new ServerManager();
                if (server != null && server.Sites != null &&
                    server.Sites.Count > 0)
                {
                    foreach (var site in server.Sites)
                    {
                        if (site.Name == ConfigurationManager.AppSettings["WebsiteName"])
                        {
                            site.Stop();

                            try
                            {
                                if (site.State == ObjectState.Stopped)
                                {
                                    //do deployment tasks...
                                    string fileName = "User.MDF";
                                    string sourcePath = ConfigurationManager.AppSettings["SourcePath"];
                                    string targetPath = ConfigurationManager.AppSettings["TargetPath"] + System.DateTime.Now.Day.ToString() + System.DateTime.Now.ToString("MMMM") + System.DateTime.Now.Year.ToString() + "\\";

                                    // Use Path class to manipulate file and directory paths.
                                    string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                    string destFile = System.IO.Path.Combine(targetPath, fileName);

                                    // To copy a folder's contents to a new location:
                                    // Create a new target folder, if necessary.
                                    if (!System.IO.Directory.Exists(targetPath))
                                    {
                                        System.IO.Directory.CreateDirectory(targetPath);
                                    }

                                    // To copy a file to another location and 
                                    // overwrite the destination file if it already exists.
                                    System.IO.File.Copy(sourceFile, destFile, true);

                                    // To copy all the files in one directory to another directory.
                                    // Get the files in the source folder. (To recursively iterate through
                                    // all subfolders under the current directory, see
                                    // "How to: Iterate Through a Directory Tree.")
                                    // Note: Check for target path was performed previously
                                    //       in this code example.
                                    if (System.IO.Directory.Exists(sourcePath))
                                    {
                                        string[] files = System.IO.Directory.GetFiles(sourcePath);

                                        // Copy the files and overwrite destination files if they already exist.
                                        foreach (string s in files)
                                        {
                                            // Use static Path methods to extract only the file name from the path.
                                            fileName = System.IO.Path.GetFileName(s);
                                            destFile = System.IO.Path.Combine(targetPath, fileName);
                                            System.IO.File.Copy(s, destFile, true);
                                        }
                                    }
                                    else
                                    {
                                        sucess = false;
                                    }
                                    // Keep console window open in debug mode.
                                }
                                else
                                {
                                    sucess = false;
                                    throw new InvalidOperationException("Could not stop website!");
                                }
                            }
                            catch
                            {
                                sucess = false;
                                throw new InvalidOperationException("Could not find website!");
                            }
                            finally
                            {
                                //restart the site...
                                site.Start();

                                string secondaryServerStatus = "";
                                var request = (HttpWebRequest)WebRequest.Create(System.Configuration.ConfigurationManager.AppSettings["SecondaryServerURL"].ToString());
                                request.Timeout = 5000;
                                try
                                {
                                    var response = (HttpWebResponse)request.GetResponse();
                                    if (response.StatusCode == HttpStatusCode.OK)
                                    {
                                        secondaryServerStatus = "Secondary license server (license2.bonexpert.com) services are up and running.";
                                    }
                                }
                                catch (WebException wex)
                                {
                                    secondaryServerStatus = "ERROR in running secondary license server (license2.bonexpert.com) services because " + wex.Message.ToString();
                                }

                                string notificationReceiver = System.Configuration.ConfigurationManager.AppSettings["NotificationReceiver"];
                                if (!string.IsNullOrEmpty(notificationReceiver))
                                {
                                    if (sucess)
                                    {
                                        SendMail(notificationReceiver, "BoneXpertSecondary@gmail.com", "", "User DB Backup", "User DB Backup sucessfully done on " + System.DateTime.Now.ToLongDateString() + ". <br/><br/> " + secondaryServerStatus);
                                    }
                                    else
                                    {
                                        SendMail(notificationReceiver, "BoneXpertSecondary@gmail.com", "", "User DB Backup", "User DB Backup failed on " + System.DateTime.Now.ToLongDateString() + ". <br/><br/> " + secondaryServerStatus);
                                    }
                                }
                            }
                        }
                    }
                }

                backupDBTimer.Interval = 60 * 60 * 24 * 1000;

            }
            catch (Exception exception)
            {

            }
        }

        private string SendMail(string toList, string from, string ccList, string subject, string body)
        {

            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient();
            string msg = string.Empty;
            try
            {
                System.Net.Mail.MailAddress fromAddress = new System.Net.Mail.MailAddress(from);
                message.From = fromAddress;
                string[] allReceivers = toList.Split(';');
                foreach (string receiver in allReceivers)
                    message.To.Add(receiver);
                if (ccList != null && ccList != string.Empty)
                    message.CC.Add(ccList);
                message.Subject = subject;
                message.IsBodyHtml = true;
                message.Body = body;
                smtpClient.Host = "smtp.gmail.com";   // We use gmail as our smtp client
                smtpClient.Port = 587;
                smtpClient.EnableSsl = true;
                smtpClient.UseDefaultCredentials = true;
                smtpClient.Credentials = new System.Net.NetworkCredential("BoneXpertSecondary@gmail.com", "Visiana2970");
                smtpClient.Send(message);
                msg = "Successful";
            }
            catch (Exception ex)
            {
                msg = ex.Message;
            }
            return msg;
        }
    }
}
