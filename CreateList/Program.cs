using Microsoft.SharePoint.Client;
using System;
using System.Security;

namespace CreateList
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("creating list");

            SharePointWorker SW = new SharePointWorker(args[0], args[1], args[2], args[3]);
            
                SW.CreateSPList();
                SW.CreateColumnsInList();
            
               
        }

    }
    class SharePointWorker {
        public string SPUrl { get; set; }
        public string ListTItle { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        private ClientContext ClientContextG = null;
        public bool CreateSPList()
        {

            logger.log("Creating Sharepoint list");
            bool status = false;

            try
            {
                Web oWebsite = ClientContextG.Web;
                ListCreationInformation listCreationInfo = new ListCreationInformation();
                listCreationInfo.Title = ListTItle;
                listCreationInfo.TemplateType = (int)ListTemplateType.GenericList; //custom list

                List oList = oWebsite.Lists.Add(listCreationInfo);
                ClientContextG.ExecuteQueryRetry();
                logger.log("list creation sucessful");
                status = true;
            }
            catch (Exception e)
            {

                logger.log("Failed to create list");
                logger.log(e.Message);
                status = false;
            }

            return status;
        }
        public bool CreateColumnsInList()
        {
            bool status = false;
            logger.log("adding columns");
            List oList = ClientContextG.Web.GetListByTitle(ListTItle);
            //-----------------------------------
            try
            {
                logger.log("creating text column \"Channel\"");

                string schemaTextField = "<Field ID='{8D294329-DD13-41A4-B817-33A2AC68F08F}' Type='Text' Name='Channel' StaticName='Channel' DisplayName='Channel' />";
                Field simpleTextField = oList.Fields.AddFieldAsXml(schemaTextField, true, AddFieldOptions.AddToDefaultContentType);
                ClientContextG.ExecuteQueryRetry();
                logger.log("Finished creating text column \"Channel\"");
                status = true;
            }
            catch (Exception e)
            {
                logger.log("failed creating text column \"Channel\"");
                logger.log(e.Message);
                status = false;
            }
            //-----------------------------------


            //-----------------------------------
            try
            {
                logger.log("creating text column \"Thread\"");
                // List oList = ClientContextG.Web.GetListByTitle(ListTItle);
                string schemaTextField = "<Field ID='{C2561E52-F544-4BF3-ADFA-1395F49043BB}' Type='Text' Name='Thread' StaticName='Thread' DisplayName='Thread' />";
                Field simpleTextField = oList.Fields.AddFieldAsXml(schemaTextField, true, AddFieldOptions.AddToDefaultContentType);
                ClientContextG.ExecuteQueryRetry();
                logger.log("Finished creating text column \"Thread\"");
                status = true;
            }
            catch (Exception e)
            {
                logger.log("failed creating text column \"Thread\"");
                logger.log(e.Message);
                status = false;
            }
            //-----------------------------------

            //-----------------------------------
            try
            {
                logger.log("creating text column \"Subject\"");
                //  List oList = ClientContextG.Web.GetListByTitle(ListTItle);
                string schemaTextField = "<Field ID='{ECA2C3E0-A774-4815-A170-803250C0E4FC}' Type='Text' Name='Subject' StaticName='Subject' DisplayName='Subject' />";
                Field simpleTextField = oList.Fields.AddFieldAsXml(schemaTextField, true, AddFieldOptions.AddToDefaultContentType);
                ClientContextG.ExecuteQueryRetry();
                logger.log("Finished creating text column \"Subject\"");
                status = true;
            }
            catch (Exception e)
            {
                logger.log("failed creating text column \"Subject\"");
                logger.log(e.Message);
                status = false;
            }
            //-----------------------------------

            //-----------------------------------
            try
            {
                logger.log("creating text column \"WebClientReadURL\"");
                //   List oList = ClientContextG.Web.GetListByTitle(ListTItle);
                string schemaUrlField = "<Field ID='{4267C5B0-59EB-4132-82C7-CA093260ABAB}' Type='URL' Name='WebClientReadURL' StaticName='WebClientReadURL' DisplayName='WebClientReadURL' Format='Hyperlink'/>";
                Field urlField = oList.Fields.AddFieldAsXml(schemaUrlField, true, AddFieldOptions.AddFieldInternalNameHint);
                ClientContextG.ExecuteQueryRetry();
                logger.log("Finished creating text column \"WebClientReadURL\"");
                status = true;
            }
            catch (Exception e)
            {
                logger.log("failed creating text column \"WebClientReadURL\"");
                logger.log(e.Message);
                status = false;
            }
            //-----------------------------------

            //-----------------------------------
            try
            {
                logger.log("creating text column \"Message\"");
                //  List oList = ClientContextG.Web.GetListByTitle(ListTItle);
                string schemaRichTextField = "<Field ID='{A1D7B648-95B6-4589-89DD-5BD59070168B}' Type='Note' Name='Message' StaticName='Message' DisplayName = 'Message' NumLines = '6' RichText = 'TRUE' RichTextMode = 'FullHtml' IsolateStyles = 'TRUE' Sortable = 'FALSE' /> ";
                Field multilineTextField = oList.Fields.AddFieldAsXml(schemaRichTextField, true, AddFieldOptions.AddFieldInternalNameHint);
                ClientContextG.ExecuteQueryRetry();
                logger.log("Finished creating text column \"Message\"");
                status = true;
            }
            catch (Exception e)
            {
                logger.log("failed creating text column \"Message\"");
                logger.log(e.Message);
                status = false;
            }
            //-----------------------------------

            logger.log("finished columns");
            return status;

        }
        public SharePointWorker(string spurl, string username, string password, string splisttitle)
        {
            SPUrl = spurl;
            ListTItle = splisttitle;
            Username = username;
            Password = password;
            ClientContextG = new ClientContext(spurl);
            ClientContextG.Credentials = new SharePointOnlineCredentials(username, ConvertToSecureString(password));
        }
        private static SecureString ConvertToSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length > 0)
            {
                foreach (var c in strPassword.ToCharArray()) secureStr.AppendChar(c);
            }
            return secureStr;

        }

    }

    class logger
    {
        public static void log(string Message)
        {

            Message = DateTime.Now + " " + Message + Environment.NewLine;
            Console.WriteLine(Message);
            string filepath = Environment.CurrentDirectory + "\\CreateListlog.log";
            System.IO.File.AppendAllText(filepath, Message);

        }
    }
}
