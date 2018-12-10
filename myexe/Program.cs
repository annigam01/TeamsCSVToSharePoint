using Microsoft.SharePoint.Client;
using System;
using System.Security;

namespace myexe
{
    class Program
    {
        static void Main(string[] args)
        {
            logger.log("Got Request. Starting writing to list");

            if (args.Length == 0)
            {
                Console.WriteLine("syntax: exe /site:this /site: that");
            }
            else
            {
                SharePointConnection SPC = new SharePointConnection(args[8], args[9], args[10], args[11]);
                Worker.work(new SharePointListItem(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7]), SPC);

            }
            logger.log("Got Request. Exiting writing to list");
        }
    }
    class SharePointListItem
    {

        private string _subject;
        public string From { get; }
        public string Team { get; }
        public string Channel { get; }
        public string Thread { get; }
        public string Subject
        {
            get
            {
                return this._subject;
            }
            private set
            {

                //goal is to extract subject ONLY if it exists.
                // if input looks like =  Collaboration Program II/Teams and Yammer communications/1537510995148 then there is no subject
                // if input looks like  = Collaboration Program II/Workspaces/1537509963231/Microsoft Teams Created by calendar month ,
                // then "Microsoft Teams Created by calendar month" is the subject

                //logic = split the string by '/', then check the last element of array if its a number, no subject, if its NOT a number then use the last array item.
                string[] input = value.Split('/');
                try
                {
                    int a;
                    string lastelement = input[input.Length - 1];

                    lastelement = lastelement.Substring(lastelement.Length - 3, 3);

                    if (int.TryParse(lastelement, out a))
                    {
                        //the last element is number
                        _subject = "No Subject";
                    }
                    else
                    {
                        //the last element is NOT number so its a subject
                        _subject = input[input.Length - 1];
                    }
                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message);
                }
            }
        }
        public string Timestamp { get; }
        public string WebClientReadURL { get; }
        public string Message { get; }

        //public SharePointConnection SharepointConnection { get; set; }

        public SharePointListItem(string from, string team, string channel, string thread, string subject, string timestamp, string webclientreadurl, string message)
        {
            logger.log("Validating Read Object...");

            this.From = from;
            this.Team = team;
            this.Channel = channel;
            this.Thread = thread;
            this.Subject = subject;
            this.Timestamp = timestamp;
            this.WebClientReadURL = webclientreadurl;
            this.Message = message;

            logger.log("Validating Read Sucessful...");
        }

    }
    class SharePointConnection
    {
        public ClientContext SPOContext { get; }
        public string ListName { get; }
        public string SiteURL { get; }

        public SharePointConnection(string Siteurl, string username, string password, string listname)
        {
            logger.log("Creating Sharepoint Object");

            try
            {
                ClientContext ctx = new ClientContext(Siteurl);
                ctx.Credentials = new SharePointOnlineCredentials(username, ConvertToSecureString(password));

                this.ListName = listname;
                this.SPOContext = ctx;
                this.SiteURL = Siteurl;
                logger.log("Creating Sharepoint Object Sucessful");
            }
            catch (Exception e)
            {
                logger.log("Creating Sharepoint Object failed");
                Console.WriteLine(e.Message);
            }
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
    class Worker
    {
        public static void work(SharePointListItem SPItem, SharePointConnection SPCon)
        {
            //we assume list exist, no checking done
            WriteToSPList(SPItem, SPCon);
        }

        private static void WriteToSPList(SharePointListItem sPItem, SharePointConnection SPCon)
        {
            logger.log("Creating Sharepoint Item");
            using (ClientContext ctx = SPCon.SPOContext)
            {
                try
                {
                    logger.log($"working on {sPItem.Team}, {sPItem.Channel}, {sPItem.Thread}, {sPItem.Subject}, {sPItem.From}, {sPItem.Timestamp}, {sPItem.WebClientReadURL}, {sPItem.Message}");
                    List oList = SPCon.SPOContext.Web.GetListByTitle(SPCon.ListName);
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = oList.AddItem(itemCreateInfo);

                    oListItem["Title"] = sPItem.Team;
                    oListItem["Channel"] = sPItem.Channel;
                    oListItem["Thread"] = sPItem.Thread;
                    oListItem["Subject"] = sPItem.Subject;
                    oListItem["Editor"] = ctx.Web.EnsureUser(sPItem.From);
                    oListItem["Modified"] = DateTime.Parse(sPItem.Timestamp);
                    oListItem["WebClientReadURL"] = sPItem.WebClientReadURL;
                    oListItem["Message"] = sPItem.Message;

                    oListItem.Update();
                    ctx.ExecuteQueryRetry();
                    logger.log("Creating Sharepoint Item sucessful");
                }
                catch (Exception e)
                {
                    logger.log("Creating Sharepoint Item Failed");
                    Console.WriteLine(e.Message);
                }
            }

        }

    }
    class logger
    {
        public static void log(string Message)
        {

            Message = DateTime.Now + " " + Message + Environment.NewLine;
            Console.WriteLine(Message);
            string filepath = Environment.CurrentDirectory + "\\CreateItemlog.log";
            System.IO.File.AppendAllText(filepath, Message);

        }
    }
}
