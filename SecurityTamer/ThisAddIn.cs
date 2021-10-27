using System;
using System.Net;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace SecurityTamer
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        Outlook.Explorers explorers;

        Outlook.Explorer ActiveExplorer;

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // TODO: attached emails that are previewed (i.e. attachment single-clicked) do not get cleaned. 
            // When double clicked, they do get cleaned via Part 1 below.

            // Parse emails when selection is changed (for old emails that haven't been cleaned)
            // Part 1, if no preview is enabled, i.e. a new window is opened
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspector_NewInspector);

            // Part 2, when emails are selected in preview pane
            explorers = this.Application.Explorers;
            // Couldn't this be just `+= Explorer_NewExplorer`?
            explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(Explorer_NewExplorer);
            foreach (Outlook.Explorer explorer in explorers)
            {
                // Not certain if this is actually necessary
                ActiveExplorer = explorer;
                explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ExplorerSelectionChange);
            }


            // This does the "new message handling"
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Item_Add);

        }

        void Item_Add(object Item)
        {
            Debug.WriteLine("Item added");
            cleanTypeSelect(Item);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        void Explorer_NewExplorer(Outlook.Explorer Explorer)
        {
            Debug.WriteLine("Explorer: " + Explorer);
            if (Explorer != null)
            {
                Explorer.SelectionChange += ExplorerSelectionChange;
                // Don't need to deal with inline responses
                // Explorer.InlineResponse += new Outlook.ExplorerEvents_10_InlineResponseEventHandler(InlineHandler);
            }
        }

        void Inspector_NewInspector(Outlook.Inspector Inspector)
        {
            Debug.WriteLine("New Inspector");
            cleanTypeSelect(Inspector.CurrentItem);
        }

        void cleanTypeSelect(object Item)
        {
            try
            {
                if (Item != null)
                {
                    if (Item is Outlook.MailItem mailItem)
                    {
                        CleanMailItem(mailItem);
                    }
                    else if (Item is Outlook.MeetingItem meetingItem)
                    {
                        CleanMeetingItem(meetingItem);
                    }
                    else if (Item is Outlook.AppointmentItem appItem)
                    {
                        CleanAppointmentItem(appItem);
                    }
                    else
                    {
                        Debug.WriteLine("Item is not a mail/meeting/appointmentItem");
                    }
                }
            }
            catch (Exception)
            {
                Debug.WriteLine("Uncaught exception in Cleaning items");
            }
            
        }

        void ExplorerSelectionChange()
        {
            Debug.WriteLine("ExplorerSelectionChange");
            ActiveExplorer = this.Application.ActiveExplorer();
            if (ActiveExplorer == null) { return; }
            Outlook.Selection selection = ActiveExplorer.Selection;

            if (selection != null)
            {
                foreach (var selected in selection)
                {
                    cleanTypeSelect(selected);

                }
            }

        }

        public static void CleanAppointmentItem(Outlook.AppointmentItem appItem)
        {
            Debug.WriteLine("Can't Cleaning AppointmentItem id: {0}, subject: {1}", appItem.EntryID, appItem.Subject);
            // appointments only use plaintext or RTF, not html (?!). The following doesn't work either, although it was mentioned online
            //var htmlBody = appItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10130102");
            try
            {
                Outlook.Inspector inspector = appItem.GetInspector;
                if (inspector != null)
                {
                    CleanRTF(inspector);
                }
            }
            catch (Exception)
            {

                Debug.WriteLine("GetInspector threw an error");
            }

        }

        public static void CleanMeetingItem(Outlook.MeetingItem meetingItem)
        {
            // meeting requests. Also just RTFs body. Normal body is set though, could convert to text only?
            //Outlook.OlBodyFormat bodyType = meetingItem.BodyFormat;
            Debug.WriteLine("Can't currently cleaning meeting id: {0}, subject: {1}", meetingItem.EntryID, meetingItem.Subject);//, Enum.GetName(typeof(Outlook.OlBodyFormat), bodyType);
            // This doesn't work for accessing HTML
            //var htmlBody = meetingItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10130102");
            try
            {
                // This always fails when reached from a "New Inspector" event
                CleanRTF(meetingItem.GetInspector);
            }
            catch (COMException)
            {
                Debug.WriteLine("Error in calling CleanRTF from CleanMeetingItem");
            }
        }

        public static void CleanRTF(Outlook.Inspector inspector)
        {
            if (inspector.IsWordMail())
            {
                Word.Document document = inspector.WordEditor;
                document.Activate();
                Word.Range range = document.Range(document.Content.Start, document.Content.End);
                // range.Text works just fine. 
                //Debug.WriteLine(range.Text);
                Word.Find findObject = range.Find;
                findObject.ClearFormatting();
                findObject.Text = "safelinks";
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = "sadlinks";
                object missing = System.Reflection.Missing.Value;
                object replaceAll = Word.WdReplace.wdReplaceAll;

                try
                {
                    // This should work, but I am getting a "Command is not available" exception. Hmm
                    // https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-search-for-and-replace-text-in-documents?view=vs-2019
                    // findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                    // document.Save();
                }
                catch (Exception)
                {
                    Debug.WriteLine("Error when executing search&replace");
                }
            }
            else
            {
                Debug.WriteLine("It is not a word mail!");
            }
        }

        public static void CleanMailItem(Outlook.MailItem mailItem)
        {
            Outlook.OlBodyFormat bodyType = mailItem.BodyFormat;
            Debug.WriteLine("Cleaning mail id: {0}, subject: {1}, type: {2}", mailItem.EntryID, mailItem.Subject, Enum.GetName(typeof(Outlook.OlBodyFormat), bodyType));
            if (mailItem.Subject != null)
            {
                if (mailItem.Subject.Contains("proposal draft for review"))
                {
                    // Useful for breakpoints
                    Debug.WriteLine("Found keyword");
                }
            }
                
            string body, newbody;
            MatchEvaluator evaluator;
            bool bodyChanged = false;
            switch (bodyType)
            {
                case Outlook.OlBodyFormat.olFormatPlain:

                    body = mailItem.Body;
                    // Debug.WriteLine(body);
                    evaluator = new MatchEvaluator(PlainURLReplacer);
                    newbody = Regex.Replace(body, @"https://[^\.]*\.safelinks\.protection\.outlook\.com/\?url=([^&]*)[^ ]+reserved=0", evaluator, RegexOptions.None, Regex.InfiniteMatchTimeout);
                    // Debug.WriteLine(newbody);
                    
                    newbody = newbody.Replace("⚠ This sender is external to UCL, please take care when accessing any links or opening attachments.\r\n\r\n\r\n", "");
                    newbody = newbody.Replace("⚠ Caution: External sender\r\n\r\n\r\n", "");
                    //Debug.WriteLine(newbody2);
                    if (body != newbody)
                    {
                        mailItem.Body = newbody;
                        bodyChanged = true;
                    }

                    break;
                case Outlook.OlBodyFormat.olFormatHTML:
                    body = mailItem.HTMLBody;
                    // In replies the div gets rewritten, so if we want to rewrite those too we need to make this a regex.
                    // Example <div class=\"\" style=\"font-family:Helvetica; font-size:14px; font-style:normal; font-variant-caps:normal; font-weight:normal; letter-spacing:normal; text-align:start; text-indent:0px; text-transform:none; white-space:normal; word-spacing:0px; text-decoration:none; background-color:rgb(255,239,213); padding:1px\">\r\n<p class=\"\" style=\"font-size:11pt; line-height:10pt; font-family:Arial,Helvetica,sans-serif\">\r\n⚠ Caution: External sender</p>\r\n</div>\r\n
                    newbody = body.Replace("<div style=\"background-color:#FFEFD5; padding:1px; \">\r\n<p style=\"font-size:11pt; line-height:10pt; font-family: 'Arial','Helvetica',sans-serif;\">\r\n⚠ This sender is external to UCL, please take care when accessing any links or opening attachments.</p>\r\n</div>\r\n<br>\r\n", "");
                    newbody = body.Replace("<div style=\"background-color:#FFEFD5; padding:1px; \">\r\n<p style=\"font-size:11pt; line-height:10pt; font-family: 'Arial','Helvetica',sans-serif;\">\r\n⚠ Caution: External sender</p>\r\n</div>\r\n<br>\r\n", "");

                    evaluator = new MatchEvaluator(HTMLURLReplacer);
                    newbody = Regex.Replace(newbody, @"href=""https://[^\.]*\.safelinks\.protection\.outlook\.com/\?url=([^&]*)[^ ]+reserved=0""([^>]*)>", evaluator, RegexOptions.None, Regex.InfiniteMatchTimeout);
                    if (body != newbody)
                    {
                        mailItem.HTMLBody = newbody;
                        bodyChanged = true;
                    }
                    break;

                case Outlook.OlBodyFormat.olFormatRichText:
                    Debug.WriteLine("Type is neither Plain or HTML, it is olFormatRichText. I have never seen such a type, so I don't know what to do :P");
                    // Not seen an email of this 
                    break;

                
                }

            if (bodyChanged & !String.IsNullOrEmpty(mailItem.EntryID))
            {
                try
                {
                    mailItem.Save();
                }
                catch (UnauthorizedAccessException)
                {
                    // Can't save, but that's ok.
                }
            }
        }

        public static string PlainURLReplacer(Match match)
        {
            Group origURL = match.Groups[1];
            string decoded = WebUtility.UrlDecode(origURL.Value);
            return decoded;
        }
        public static string HTMLURLReplacer(Match match)
        {
            Group origURL = match.Groups[1];
            string decoded = WebUtility.UrlDecode(origURL.Value);
            string otherStuff = match.Groups[2].Value;
            otherStuff = Regex.Replace(otherStuff, @" (?:originalsrc|shash)=""[^""]*""", "", RegexOptions.None, Regex.InfiniteMatchTimeout);
            return String.Format("href=\"{0}\"{1}>", decoded, otherStuff);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
