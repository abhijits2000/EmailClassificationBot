using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using SE.Core.Automation.Interfaces;
using SE.Core.Automation.Models.Common;
using Email = Microsoft.Office.Interop.Outlook;
using System.IO;



namespace EmailClassificationMicrobot
{
    public class EmailClassification : IAECodePackage
    {
        #region Declaration Section
        public event EventHandler ExecutionCompleted;

        public string strJSONFilePath = string.Empty;
        public string strCSVFilePath = string.Empty;
        public string strTrainorTest = string.Empty; //"Train
        public string strEmailFolderPath = string.Empty;
        public string strPythonFilePath = string.Empty;
        public string strPythonEXEPath = string.Empty;
        public string strEmailType = string.Empty;
        public string strMessage = string.Empty;
        public string strErrDesc = string.Empty;
        public int strErrCode = 0;


        //string strEmailFolderPath = "Processed";
        //string strPythonFilePath = @"C:\Users\anuradha.agrawal\AppData\Local\Continuum\anaconda3\text.py";//string.Empty;
        //string strPythonEXEPath = @"C:\Users\anuradha.agrawal\AppData\Local\Continuum\anaconda3\python.exe";
        //string strEmailType = string.Empty;

        #endregion

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string CSVFilePath { get { return strCSVFilePath; } set { strCSVFilePath = value; } }

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string JSONFilePath { get { return strJSONFilePath; } set { strJSONFilePath = value; } }

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string TrainorTest { get { return strTrainorTest; } set { strTrainorTest = value; } }

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string EMailFolder { get { return strEmailFolderPath; } set { strEmailFolderPath = value; } }

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string PythonFilePath { get { return strPythonFilePath; } set { strPythonFilePath = value; } }

        [ArgumentDirection(Direction = DirectionType.Input)]
        public string PythonEXEPath { get { return strPythonEXEPath; } set { strPythonEXEPath = value; } }

        [ArgumentDirection(Direction = DirectionType.Output)]
        public string EmailType { get { return strEmailType; } set { strEmailType = value; } }

        [ArgumentDirection(Direction = DirectionType.Output)]
        public string Message { get { return strMessage; } set { strMessage = value; } }

        [ArgumentDirection(Direction = DirectionType.Output)]
        public string ErrDesc { get { return strErrDesc; } set { strErrDesc = value; } }

        [ArgumentDirection(Direction = DirectionType.Output)]
        public int ErrCode { get { return strErrCode; } set { strErrCode = value; } }



        public void Execute(StudioContext context)
        {
            try
            {
                if (CheckFilePaths())
                {
                    GetEmail();
                    Message = "Inside Execute() : Execution Completed Successfully.";
                }                   
                else
                    Message = "Inside Execute() Exception: File not found.";
            }
            catch (Exception ex)
            {
                ErrCode = 444;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside Execute() Exception: Error While executing Request. " + ex.Message;
                //Console.WriteLine(ex.ToString());
                throw;
            }
            ExecutionCompleted(this, null);
            // throw new NotImplementedException();
        }

        private bool CheckFilePaths()
        {
            bool IsFilePathFound = false;
            try
            {
                if (File.Exists(CSVFilePath) && File.Exists(strJSONFilePath) && File.Exists(PythonEXEPath) && File.Exists(PythonFilePath))
                    IsFilePathFound = true;
                else
                    IsFilePathFound = false;

            }
            catch (Exception ex)
            {
                ErrCode = 404;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;
                Message = "Inside CheckFilePaths() Exception: Error while checking file exists. " + ex.Message;
                //throw;
            }
            return IsFilePathFound;
        }
        private void GetEmail()
        {
            Email.Application app = null;
            Email._NameSpace ns = null;
            Email.MailItem item = null;
            Email.MAPIFolder inboxFolder = null;
            Email.MAPIFolder subFolder = null;
            Email.MAPIFolder NewIncident = null;
            Email.MAPIFolder FollowUp = null;
            Email.MAPIFolder None = null;
            string EmailType = string.Empty;

            try
            {
                app = new Email.Application();
                ns = app.GetNamespace("MAPI");
                ns.Logon(null, null, false, false);

                inboxFolder = ns.GetDefaultFolder(Email.OlDefaultFolders.olFolderInbox);
                subFolder = inboxFolder.Folders[EMailFolder]; //folder.Folders[1]; also works
                NewIncident = inboxFolder.Folders["NewIncident"];
                FollowUp = inboxFolder.Folders["FollowUp"];
                None = inboxFolder.Folders["None"];
                //Console.WriteLine("Folder Name: {0}, EntryId: {1}", subFolder.Name, subFolder.EntryID);
                //Console.WriteLine("Num Items: {0}", subFolder.Items.Count.ToString());
                if (subFolder.Items.Count != 0)
                {


                    for (int i = 1; i <= subFolder.Items.Count; i++)
                    {

                        item = (Email.MailItem)subFolder.Items[i];
                        string Body = item.Body;
                        if (TrainorTest.ToLower().Equals("train"))
                        {

                            CreateDS(Body);

                        }
                        else
                        {
                            EmailType = ClassifyMail(Body);
                            if (EmailType.ToLower().Equals("newincident"))
                            {
                                item.Move(NewIncident);
                            }
                            else if (EmailType.ToLower().Equals("followup"))
                            {
                                item.Move(FollowUp);
                            }
                            else
                            {
                                item.Move(None);
                            }
                        }
                    }

                }
                else
                {
                    ErrCode = 204;
                    ErrDesc += "Message: No Email found in " + EMailFolder + " Folder." + Environment.NewLine;

                    Message = "No Email found in " + EMailFolder + " Folder.";
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ErrCode = 503;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside GetEmail() Exception: Error While Getting Email Form Outlook. " + ex.Message;
                //Console.WriteLine(ex.ToString());
            }
            finally
            {
                ns = null;
                app = null;
                inboxFolder = null;
            }
        }

        private void CreateDS(string Body)
        {
            try
            {
                if (File.Exists(CSVFilePath))
                    File.Delete(CSVFilePath);

                if (CreateMailCSVFile(Body))
                {
                    string fileName = string.Format(PythonFilePath + " \"" + "{0}" + " \"" + " \"" + "{1}" + " \"" + " \"" + "{2}" + " \"", JSONFilePath, CSVFilePath, string.Empty);
                    run_Python(fileName);
                    Message = "Successfully created Dataset.";
                }
                else
                    Message = "Error While creating Dataset.";

            }
            catch (Exception ex)
            {
                ErrCode = 205;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside CreateDS() Exception: Error While Creating DataSet. " + ex.Message;
                // throw;
            }
        }

        private string ClassifyMail(string Body)
        {
            string strmailtype = "None";
            try
            {
                if (Getpostion(Body, "From:") != -1)
                {
                    EmailType = CreateRequest(Body);
                    strmailtype = EmailType.Replace("\r\n", string.Empty);
                    //Console.WriteLine(EmailType);
                }
                else
                {
                    string fileName = string.Format(PythonFilePath + " \"" + "{0}" + " \"" + " \"" + "{1}" + " \"" + " \"" + "{2}" + " \"", JSONFilePath, CSVFilePath, Body);
                    EmailType = run_Python(fileName);
                    strmailtype = EmailType.Replace("\r\n", string.Empty);
                    //Console.WriteLine(EmailType);

                }               
            }
            catch (Exception ex)
            {
                ErrCode = 206;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside ClassifyMail() Exception: Error While Classifying Email. " + ex.Message;
                throw;
            }
            return strmailtype;
        }

        private int Getpostion(string message, string delimeter)
        {
            int delimeterPosition = 0;
            int position = 0;

            try
            {
                delimeterPosition = message.IndexOf(delimeter);
                position = (delimeterPosition > -1) ? delimeterPosition : -1;

            }
            catch (Exception ex)
            {
                ErrCode = 400;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside Getpostion() Exception: Error While Getting Position of String in Email Content " + ex.Message;
                //Console.WriteLine(ex.ToString());
                //throw;
            }
            return position;
        }

        private string CreateRequest(string message)
        {
            int NewIncCount = 0;
            int FollowUpCount = 0;
            int NoneCount = 0;
            string strFinalmsg = string.Empty;
            string strClassification = string.Empty;
            string typeofEmail = "None";

            int fromindex = 0;            
            int subindex = -1;
            int subendindex = -1;
            string fileName = string.Empty;
            try
            {


                string currentmsg = string.Empty;
                currentmsg = message;
                do
                {

                    fromindex = Getpostion(currentmsg, "From:");


                    if (fromindex != -1)
                    {
                        strFinalmsg = currentmsg.Substring(0, fromindex);
                        fileName = string.Format(PythonFilePath + " \"" + "{0}" + " \"" + " \"" + "{1}" + " \"" + " \"" + "{2}" + " \"", JSONFilePath, CSVFilePath, strFinalmsg);
                        strClassification = run_Python(fileName);
                        if (strClassification.ToLower().Equals("Newincident"))
                            NewIncCount++;
                        else if (strClassification.ToLower().Equals("FollowUp"))
                            FollowUpCount++;
                        else
                            NoneCount++;

                        subindex = Getpostion(currentmsg, "Subject:");
                        currentmsg = currentmsg.Substring(subindex + 1, (currentmsg.Length - (subindex + 1)));


                        subendindex = Getpostion(currentmsg, "\r\n");
                        currentmsg = currentmsg.Substring(subendindex + 1, (currentmsg.Length - (subendindex + 1)));
                    }
                    else
                    {
                        fileName = string.Format(PythonFilePath + " \"" + "{0}" + " \"" + " \"" + "{1}" + " \"" + " \"" + "{2}" + " \"", JSONFilePath, CSVFilePath, currentmsg);
                        strClassification = run_Python(fileName);
                        if (strClassification.ToLower().Equals("newincident"))
                            NewIncCount++;
                        else if (strClassification.ToLower().Equals("followup"))
                            FollowUpCount++;
                        else
                            NoneCount++;
                    }
                } while (fromindex != -1);


                if (NewIncCount > FollowUpCount)
                    typeofEmail = "Newincident";
                else if (FollowUpCount > NewIncCount)
                    typeofEmail = "FollowUp";
                else
                    typeofEmail = "None";
            }
            catch (Exception ex)
            {
                ErrCode = 401;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                Message = "Inside CreateRequest() Exception: Error While Classifying Email. " + ex.Message;
                //Console.WriteLine(ex.ToString());

            }
            return typeofEmail;
        }

        private string run_Python(string PythonInput)
        {
            string output = string.Empty;
            Process p = new Process();
            try
            {
                string fileName = string.Format(PythonInput);

                

                p.StartInfo.FileName = PythonEXEPath; //Python.exe location
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.UseShellExecute = false; // ensures you can read stdout
                p.StartInfo.Arguments = fileName; // start the python program with two parameters
                p.Start(); // start the process (the python program)                          
                StreamReader s = p.StandardOutput;

               // Thread.Sleep(10000);
                output = s.ReadToEnd();
               // Console.WriteLine(output);
               // p.WaitForExit();


            }
            catch (Exception ex)
            {
                ErrCode = 420;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;


                Message = "Inside run_Python() Exception: Error While Executing Python Script. " + ex.Message;
                //Console.WriteLine(ex.ToString());
            }
            finally
            {
               // p.Close();
                //p.Dispose();
            }

            return output;
        }

        private bool CreateMailCSVFile(string message)
        {

            string strFinalmsg = string.Empty;
            string strClassification = string.Empty;
            bool isFileCreated = false;

            int fromindex = 0;

            int subindex = -1;
            int subendindex = -1;
            try
            {


                string currentmsg = string.Empty;
                currentmsg = message;

                do
                {

                    fromindex = Getpostion(currentmsg, "From:");


                    if (fromindex != -1)
                    {
                        strFinalmsg = currentmsg.Substring(0, fromindex);




                        using (StreamWriter sw = new StreamWriter(CSVFilePath, true))
                        {
                            sw.WriteLine(strFinalmsg);
                        }



                        subindex = Getpostion(currentmsg, "Subject:");
                        currentmsg = currentmsg.Substring(subindex + 1, (currentmsg.Length - (subindex + 1)));


                        subendindex = Getpostion(currentmsg, "\r\n");
                        currentmsg = currentmsg.Substring(subendindex + 1, (currentmsg.Length - (subendindex + 1)));
                    }
                    else
                    {
                        using (StreamWriter sw = new StreamWriter(CSVFilePath, true))
                        {
                            sw.WriteLine(currentmsg);
                        }

                    }
                } while (fromindex != -1);

                isFileCreated = true;

            }
            catch (Exception ex)
            {
                isFileCreated = false;
                ErrCode = 402;
                ErrDesc += "Message: " + ex.Message + Environment.NewLine + "Inner Exception: " + ex.InnerException + Environment.NewLine;

                
                Message = "Inside CreateMailCSVFile() Exception: Error While creating Dataset. " + ex.Message;
                //Console.WriteLine(ex.ToString());
            }
            return isFileCreated;
        }

       
    }
}
