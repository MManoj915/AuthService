using Dapper;
using MimeKit;
using MailKit.Net.Smtp;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using WinSCP;
using System.Net;
using System.Xml;
using PBMAuthorizationService;
using PBMAuthorizationService.PBMSwitchReference;
using PBMAuthorizationService.DHPOReference;

namespace PBMAuthorizationService
{
    public partial class Service1 : ServiceBase
    {
        private Thread _thread;
        private int ThreadTimeLimit = 1;
        private static IDbConnection _db = new OracleConnection(ConfigurationSettings.AppSettings["ConnectionString"].ToString());
        PBMSwitchReference.PayerIntegrationWSService PBMsrv;
        public Service1()
        {
            string TimeLimit = string.Empty;
            int HourLimit = Convert.ToInt32(ConfigurationSettings.AppSettings["HOURLIMIT"].ToString());
            if (string.IsNullOrEmpty(TimeLimit))
                TimeLimit = "24";
            ThreadTimeLimit = (Convert.ToInt16(TimeLimit) * HourLimit);
            InitializeComponent();
            //Execute();
        }

        protected override void OnStart(string[] args)
        {
            this._thread = new Thread(new ThreadStart(this.Execute));
            this._thread.Start();
        }

        private void Execute1()
        {

            List<string> Errors = new List<string>();

            bool doprocess = false;
            try
            {
                while (!doprocess)
                {
                    List<AuthBatchID> BatchList = _db.Query<AuthBatchID>(" Select * from IM_AUTHBATCH").ToList();
                    PBMAuthorizationService.PBMSwitchReference.pbmPriorAuthorization[] files;
                    for (int i = 0; i < BatchList.Count; i++)
                    {
                        transactionBatch t = PBMsrv.findTransactionsByBatchID("A025", "8e260394-2211-4a4a-b14b-7f0824183f15", BatchList[i].BatchID, 2);
                        priorAuthorizationBatch response = t.priorAuthorizationBatch;
                        files = response.authorizationsList;
                        foreach (pbmPriorAuthorization file in files)
                        {
                            try
                            {
                                string XMLData = GenaratePBMPriorAuthorizationXmlFile(file);
                                string FileLocation = WriteAuthorizationPBMSubmissionFile(file.priorAuthorization.Header.TransactionDate, file.priorAuthorization.Header.ReceiverID, file.priorAuthorization.Header.ReceiverID, XMLData); //Create the downloaded file in the file system
                                WritePriorAuthorizationRequestFilesLogToDB(response.batchID, file, 1, 1, "", 3, FileLocation);// log file to DB

                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                    }

                    Thread.Sleep(ThreadTimeLimit);
                }
            }
            catch (Exception EX)
            {

            }
        }

        private void Execute()
        {

            List<string> Errors = new List<string>();

            bool doprocess = false;
            try
            {
                while (!doprocess)
                {
                    List<AuthBatchID> BatchList = _db.Query<AuthBatchID>(" Select * from IM_AUTHBATCH").ToList();
                    PBMAuthorizationService.PBMSwitchReference.pbmPriorAuthorization[] files;
                    for (int i = 0; i < BatchList.Count; i++)
                    {
                        transactionBatch t = PBMsrv.findTransactionsByBatchID("A025", "8e260394-2211-4a4a-b14b-7f0824183f15", BatchList[i].BatchID, 2);
                        priorAuthorizationBatch response = t.priorAuthorizationBatch;
                        files = response.authorizationsList;
                        foreach (pbmPriorAuthorization file in files)
                        { 
                            try
                            {
                                string XMLData = GenaratePBMPriorAuthorizationXmlFile(file);
                                string FileLocation = WriteAuthorizationPBMSubmissionFile(file.priorAuthorization.Header.TransactionDate, file.priorAuthorization.Header.ReceiverID, file.priorAuthorization.Header.ReceiverID, XMLData); //Create the downloaded file in the file system
                                WritePriorAuthorizationRequestFilesLogToDB(response.batchID, file, 1, 1, "", 3, FileLocation);// log file to DB
                                     
                            }
                            catch (Exception ex)
                            { 
                                continue;
                            } 
                        } 
                    }

                    Thread.Sleep(ThreadTimeLimit);
                }
            }
            catch (Exception EX)
            {

            }
        }

        protected override void OnStop()
        {
            if (this._thread != null)
            {
                this._thread.Abort();
                this._thread.Join();
            }
        } 

        public string GenaratePBMPriorAuthorizationXmlFile(pbmPriorAuthorization ClaimFile)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder builder = new StringBuilder();
            builder.Append("<Prior.Authorization> \n\t");
            builder.Append("<Header>\n\t\t");
            builder.Append("<SenderID>" + ClaimFile.priorAuthorization.Header.SenderID + "</SenderID>\n\t\t");
            builder.Append("<ReceiverID>" + ClaimFile.priorAuthorization.Header.ReceiverID + "</ReceiverID>\n\t\t");
            builder.Append("<TransactionDate>" + ClaimFile.priorAuthorization.Header.TransactionDate + "</TransactionDate>\n\t\t");
            builder.Append("<RecordCount>" + ClaimFile.priorAuthorization.Header.RecordCount + "</RecordCount>\n\t\t");
            builder.Append("<DispositionFlag>" + ClaimFile.priorAuthorization.Header.DispositionFlag + "</DispositionFlag>\n\t");
            builder.Append("</Header>\n\t");
            builder.Append("<Authorization>\n\t\t");
            builder.Append("<Result>" + ClaimFile.priorAuthorization.Authorization.Result + "</Result>\n\t\t");
            builder.Append("<ID>" + ClaimFile.priorAuthorization.Authorization.ID + "</ID>\n\t\t");
            builder.Append("<IDPayer>" + ClaimFile.priorAuthorization.Authorization.IDPayer + "</IDPayer>\n\t\t");
            builder.Append("<DenialCode>" + ClaimFile.priorAuthorization.Authorization.DenialCode + "</DenialCode>\n\t\t");
            builder.Append("<Start>" + ClaimFile.priorAuthorization.Authorization.Start + "</Start>\n\t\t");
            builder.Append("<End>" + ClaimFile.priorAuthorization.Authorization.End + "</End>\n\t\t");
            builder.Append("<Limit>" + ClaimFile.priorAuthorization.Authorization.Limit + "</Limit>\n\t\t");
            builder.Append("<Comments>" + ClaimFile.priorAuthorization.Authorization.Comments + "</Comments>\n");
            for (int l = 0; l < ClaimFile.priorAuthorization.Authorization.Activity.Count(); l++)
            {
                builder.Append("<Activity>\n\t\t\t");
                builder.Append("<ID>" + ClaimFile.priorAuthorization.Authorization.Activity[l].ID + "</ID>\n\t\t\t");
                builder.Append("<Type>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Type + "</Type>\n\t\t\t");
                builder.Append("<Code>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Code + "</Code>\n\t\t\t");
                builder.Append("<Quantity>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Quantity + "</Quantity>\n\t\t\t");
                builder.Append("<Net>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Net + "</Net>\n\t\t\t");
                builder.Append("<List>" + ClaimFile.priorAuthorization.Authorization.Activity[l].List + "</List>\n\t\t\t");
                builder.Append("<PatientShare>" + ClaimFile.priorAuthorization.Authorization.Activity[l].PatientShare + "</PatientShare>\n\t\t\t");
                builder.Append("<PaymentAmount>" + ClaimFile.priorAuthorization.Authorization.Activity[l].PaymentAmount + "</PaymentAmount>\n\t\t\t");
                builder.Append("<DenialCode>" + ClaimFile.priorAuthorization.Authorization.Activity[l].DenialCode + "</DenialCode>\n\t\t\t");

                if (ClaimFile.priorAuthorization.Authorization.Activity[l].Observation != null)
                {
                    for (int m = 0; m < ClaimFile.priorAuthorization.Authorization.Activity[l].Observation.Count(); m++)
                    {
                        builder.Append("<Activity_Observation>\n\t\t\t\t");
                        builder.Append("<Type>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Observation[m].Type + "</Type>\n\t\t\t\t");
                        builder.Append("<Code>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Observation[m].Code + "</Code>\n\t\t\t\t");
                        builder.Append("<Value>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Observation[m].Value + "</Value>\n\t\t\t\t");
                        builder.Append("<ValueType>" + ClaimFile.priorAuthorization.Authorization.Activity[l].Observation[m].ValueType + "</ValueType>\n\t\t\t");
                        builder.Append("</Activity_Observation>\n\t\t");
                    }
                }
                builder.Append("</Activity>\n\t");
            }
            builder.Append("</Authorization>\n\t");
            builder.Append("</Prior.Authorization>\n\t");
            sb.Append(builder.ToString());
            return sb.ToString();
        }

        public string WriteAuthorizationPBMSubmissionFile(string TransactionDate, string SenderID, string fileName, string XMLData)
        {
            string CurrentDateTime = DateTime.Now.ToString("yyyy/MM/dd");
            CurrentDateTime = CurrentDateTime.Replace("/", "");
            string path = "D:\\Eclaims\\PriorAuthorization\\" + CurrentDateTime + "\\" + SenderID; 
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            bool Existed = false;
            int count = 1;
            string FileName = path + "\\" + fileName + "_" + CurrentDateTime;
            string tmpFileName = FileName + ".xml";
            while (File.Exists(tmpFileName))
            {
                Existed = true;
                tmpFileName = FileName + "-" + count + ".xml";
                count++;
            }
            if (!Existed)
                FileName = FileName + ".xml";
            else
                FileName = tmpFileName;

            File.WriteAllText(FileName, XMLData);

            return FileName;
        }

        public void WritePriorAuthorizationRequestFilesLogToDB(string BatchID, PBMSwitchReference.pbmPriorAuthorization file, int TransactionType, int TransactionStatus, string TransactionError, int SYS_FILE_SOURCE, string FileLocation)
        {
            string Pkey = "select  nvl(max(SYS_ID),0)+1 from AUTH_BATCH_FILES";
            int Sys_ID = 0;
            Sys_ID = Convert.ToInt32(_db.ExecuteScalar(Pkey));
            string sqlcmd = " INSERT INTO AUTH_BATCH_FILES (SYS_ID,FILE_ID, FILE_NAME, SENDER_ID, " +
                " RECEIVER_ID, TRANSACTION_DATE, RECORD_COUNT, TRANSACTION_TYPE, TRANSACTION_ERROR, " +
                " FILE_LOCATION, STATUS, SYS_FILE_SOURCE) " +
                " VALUES ("+ Sys_ID + ",'"+ BatchID + "', '"+file.fileName+"', '"+ file.priorAuthorization.Header.SenderID + "', " +
                " '"+ file.priorAuthorization.Header.ReceiverID + "', '"+ file.priorAuthorization.Header.TransactionDate + "', " +
                " '"+ file.priorAuthorization.Header.RecordCount + "', "+ TransactionType + ", null, '"+ FileLocation + "', "+TransactionStatus+",3)  ";

            _db.Execute(sqlcmd);
        }
    }


}
