using NotifySecurity.Properties;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text;

namespace NotifySecurity
{
    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;
        public Ribbon1()
        {
            StartUp = true;
        }

        public Boolean StartUp = false;
        public String ddlEntityValue = "Company";
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("NotifySecurity.Ribbon1.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public string GetContextMenuLabel(IRibbonControl control)
        {
            return "Smarthouse Security";
        }

        public string GetGroupLabel(IRibbonControl control)
        {
            return "";
        }

        public string GetTabLabel(IRibbonControl control)
        {
            return "Smarthouse Security";
        }

        public string GetSupertipLabel(IRibbonControl control)
        {
            var v = Assembly.GetAssembly(typeof(Ribbon1)).GetName().Version;
            int revMaj = v.Major;
            int revMin = v.Minor;
            int revBuild = v.Build;
            int revRev = v.Revision;
            //return "Shieldy v" + revMaj.ToString() + "." + revMin.ToString() + "." + revBuild.ToString() + "." + revRev.ToString();// versionInfo.ToString();
            return "";
        }

        public string GetScreentipLabel(IRibbonControl control)
        {
            return "Smarthouse Security";
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            return "Smarthouse Security";
        }

        public void ShowMessageClick(IRibbonControl control)
        {

            CreateNewMailToSecurityTeam(control);
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            return new Bitmap(Resources.sh);
        }

        private void CreateNewMailToSecurityTeam(IRibbonControl control)
        {
            Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            if (selection.Count == 1)   // Check that selection is not empty.
            {
                object selectedItem = selection[1];   // Index is one-based.
                Object mailItemObj = selectedItem as Object;
                MailItem mailItem = null;// selectedItem as MailItem;
                if (selection[1] is MailItem)
                {
                    mailItem = selectedItem as MailItem;
                }
                MailItem tosend = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                tosend.Attachments.Add(mailItemObj);
                try
                {
                    tosend.To = Settings.Default.Security_Team_Mail;
                    tosend.Subject = "[User Alert] Suspicious mail received. Please investigate";
                    tosend.CC = Settings.Default.Security_Team_Mail_cc;
                    tosend.BCC = Settings.Default.Security_Team_Mail_bcc;

                    string allHeaders = "";
                    if (selection[1] is MailItem)
                    {
                        string[] preparedByArray = mailItem.Headers("X-PreparedBy");
                        string preparedBy;
                        if (preparedByArray.Length == 1)
                        {
                            preparedBy = preparedByArray[0];
                        }
                        else
                        {
                            preparedBy = "";
                        }
                        allHeaders = mailItem.HeaderString();
                    }
                    else
                    {
                        string typeFound = "unknown";
                        typeFound = (selection[1] is MailItem) ? "MailItem" : typeFound;
                        if (typeFound == "unknown")
                        {
                            typeFound = (selection[1] is MeetingItem) ? "MeetingItem" : typeFound;
                        }   
                        if (typeFound == "unknown")
                        {
                            typeFound = (selection[1] is ContactItem) ? "ContactItem" : typeFound;
                        }
                        if (typeFound == "unknown")
                        {
                            typeFound = (selection[1] is AppointmentItem) ? "AppointmentItem" : typeFound;
                        }
                        if (typeFound == "unknown")
                        {
                            typeFound = (selection[1] is TaskItem) ? "TaskItem" : typeFound;
                        }
                        allHeaders = "Selected Outlook item was not a mail (" + typeFound + "), no header extracted";
                    }
                    string SwordPhishURL = SwordphishObject.SetHeaderIDtoURL(allHeaders);
                    if (SwordPhishURL != SwordphishObject.NoHeaderFound)
                    {
                        string SwordPhishAnswer = SwordphishObject.SendNotification(SwordPhishURL);
                    }
                    else
                    {
                        StringBuilder BodyContent = new StringBuilder("Hello, I received the attached email and I think it is suspicious");
                        BodyContent.AppendLine();
                        BodyContent.Append("I think this mail is malicious for the following reasons:");
                        BodyContent.AppendLine();
                        BodyContent.Append("Please analyze and provide some feedback.");
                        BodyContent.AppendLine();
                        BodyContent.AppendLine();
                        BodyContent.Append(GetCurrentUserInfos());
                        BodyContent.AppendLine();
                        BodyContent.AppendLine();
                        BodyContent.Append("Message headers:");
                        BodyContent.AppendLine();
                        BodyContent.Append("--------------");
                        BodyContent.AppendLine();
                        BodyContent.Append(allHeaders);
                        BodyContent.AppendLine();
                        BodyContent.AppendLine();
                        tosend.Body = BodyContent.ToString();
                        tosend.Save();
                        tosend.Display();
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Using default template" + ex.Message);
                    MailItem mi = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    mi.To = Settings.Default.Security_Team_Mail;
                    mi.Subject = "Security addin error";
                    String txt = ("An error occured, please notify your security contact and give him/her the following information: " + ex);
                    mi.Body = txt;
                    mi.Save();
                    mi.Display();
                }
            }
            else if (selection.Count < 1)   // Check that selection is not empty.
            {
                MessageBox.Show("Please select one mail.");
            }
            else if (selection.Count > 1)
            {
                MessageBox.Show("Please select only one mail to be raised to the security team.");
            }
            else
            {
                MessageBox.Show("Bad luck... this case has not been identified by the dev");
            }
        }

        public String GetCurrentUserInfos()
        {
            StringBuilder wComputername = new StringBuilder(string.Format("{0} ({1})", Environment.MachineName, Environment.OSVersion.ToString()));
            StringBuilder wUsername = new StringBuilder(string.Format("{0}\\{1}", Environment.UserDomainName, Environment.UserName));
            StringBuilder usefullInfo = new StringBuilder("Possibly useful information:\n--------------");
            AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                ExchangeUser currentUser = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    usefullInfo.AppendLine();
                    usefullInfo.Append(string.Format("Name: {0}", currentUser.Name));
                    usefullInfo.AppendLine();
                    usefullInfo.Append(string.Format("STMP address: {0}", currentUser.PrimarySmtpAddress));
                    usefullInfo.AppendLine();
                    usefullInfo.Append(string.Format("Business phone: {0}", currentUser.BusinessTelephoneNumber));
                    usefullInfo.AppendLine();
                    usefullInfo.Append(string.Format("Mobile phone: {0}", currentUser.MobileTelephoneNumber));
                }
            }
            usefullInfo.AppendLine();
            usefullInfo.Append(string.Format("Windows username: {0}", wUsername.ToString()));
            usefullInfo.AppendLine();
            usefullInfo.Append(string.Format("Computername: {0}", wComputername.ToString()));
            usefullInfo.AppendLine();
            return usefullInfo.ToString();
        }
    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
            {
                return headers[name].ToArray();
            }
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches(headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(h => h.Groups["header_key"].Value, h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor.GetProperty(TransportMessageHeadersSchema);
        }

    }
}