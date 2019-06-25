using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.ServiceModel.Syndication;
using System.Runtime.InteropServices;
using System.Xml;
using System.Security.Permissions;

namespace NSMcGeeDashboard
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]                                                      //Enable invoking of Javascript functions

    public partial class McGeeDashboard : Form
    {
        // Global variables, needed to pass to multiple functions
        HtmlElement mobilengineContent;
        bool mobEngReset;
        //HtmlElement mobilengineHeader;

        List<String> FeedSubject = new List<String>();
        List<String> FeedSummary = new List<String>();

        HtmlDocument editorDoc;
        string editorDocPath;
        string editorDocText;

        int safetyLadderScrollPos = 0;
        int rssFeedCounter = 1;
        int readyBrowsers = 0;
        bool editorActive = false;
        bool editorEnlarged = false;
        
        public static IEnumerable<HtmlElement> ElementsByClass(HtmlDocument doc, string className)                  //gets a collection of HtmlElements with a particular class name from the specified HtmlDocument
        {
            foreach (HtmlElement e in doc.All)
                if (e.GetAttribute("className") == className)
                    yield return e;
        }

        public static IEnumerable<HtmlElement> ElementsById(HtmlDocument doc, string Id)                            //gets a collection of HtmlElements with a particular id from the specified HtmlDocument
        {
            foreach (HtmlElement e in doc.All)
                if (e.GetAttribute("Id") == Id)
                    yield return e;
        }

        public static IEnumerable<HtmlElement> ElementsByTag(HtmlDocument doc, string tag)                          //gets a collection of HtmlElements with a particular tag from the specified HtmlDocument
        {
            foreach (HtmlElement element in doc.GetElementsByTagName(tag))
                yield return element;
        }

        public McGeeDashboard()                                                     //constructor for the McGee Dashboard class
        {
            InitializeSplashScreen();

            RegistryKey myKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);                  //sets WebBrowser controls to emulate Internet Explorer 11
            if (myKey != null)
            {
                myKey.SetValue(System.AppDomain.CurrentDomain.FriendlyName, "11000", RegistryValueKind.DWord);
                myKey.Close();
            }

            InitializeComponent();                                                  //adds all the controls to the form

            mobilengine_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\mobresource\\Dashboard.html");
            mobilengine_WebBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(mobilengine_WebBrowser_DocumentCompleted);

            dashboardTitle_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Dashboard Title.html");
            dashboardTitle_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            mandatoryRequirement_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Mandatory Requirement.html");
            mandatoryRequirement_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            focusBoard_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Focus Board.html");
            focusBoard_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            siteName_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Site Name.html");
            siteName_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            siteInformation_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Site Information.html");
            siteInformation_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            bulletin_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Bulletin.html");
            bulletin_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            twitterTitle_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Twitter Title.html");
            twitterTitle_WebBrowser.Document.Click += new HtmlElementEventHandler(WebBrowser_DocumentClicked);

            mcgeeEditor_WebBrowser.Url = new Uri(Application.StartupPath + "\\Resources\\Editor.html");
            mcgeeEditor_WebBrowser.ObjectForScripting = this;                                 //allow the WebBrowser document to call functions within this class

            Timer dateTimer = new Timer();                                          //creates a new Timer class to update the Time and Date
            dateTimer.Interval = 1000;
            dateTimer.Tick += DateTimer_Tick;
            dateTimer.Start();

            currentDateTime_RichTextBox.SelectionAlignment = HorizontalAlignment.Center;           //set the control's text to horizontally aligned

            Timer mobilengine_Timer = new Timer();                                   //creates a new instance of the Timer class to scroll Mobilengine
            mobilengine_Timer.Interval = 5000;
            mobilengine_Timer.Tick += Mobilengine_Timer_Tick;

            Timer safetyLadder_Timer = new Timer();                                  //creates a new instance of the Timer class to scroll the safety ladder
            safetyLadder_Timer.Interval = 5000;
            safetyLadder_Timer.Tick += SafetyLadder_Timer_Tick;

            Timer rssFeed_Timer = new Timer();                                       //creates a new instance of the Timer class to update the RssFeed
            rssFeed_Timer.Interval = 10000;
            rssFeed_Timer.Tick += RssFeed_Timer_Tick;

            Timer refreshDashboard_Timer = new Timer();
            refreshDashboard_Timer.Interval = (int)TimeSpan.FromHours(1).TotalMilliseconds;
            refreshDashboard_Timer.Tick += RefreshDashboard_Timer_Tick;

            safetyLadder_Panel.VerticalScroll.Enabled = true;
            safetyLadder_Panel.VerticalScroll.Visible = true;
            safetyLadder_Panel.VerticalScroll.Minimum = 0;
            safetyLadder_Panel.VerticalScroll.Maximum = safetyLadder_PictureBox.Height;

            XmlReader skyNews_Reader = XmlReader.Create("http://feeds.skynews.com/feeds/rss/uk.xml");
            SyndicationFeed feed = SyndicationFeed.Load(skyNews_Reader);

            foreach (SyndicationItem item in feed.Items)
            {
                String subject;
                String summary;

                if (item.Title.Text == String.Empty)
                    subject = "No Subject";                         //account for empty value
                else
                {
                    subject = System.Net.WebUtility.HtmlDecode(item.Title.Text);                                            //decode html characters
                    subject = System.Text.RegularExpressions.Regex.Replace(subject, "<a href='.*'>", String.Empty);         //remove hyperlinks
                    subject = System.Text.RegularExpressions.Regex.Replace(subject, "</a>", String.Empty);
                }

                if (item.Summary.Text == String.Empty)
                    summary = "No Summary";
                else
                {
                    summary = System.Net.WebUtility.HtmlDecode(item.Summary.Text);
                    summary = System.Text.RegularExpressions.Regex.Replace(summary, "<a href='.*'>", String.Empty);
                    summary = System.Text.RegularExpressions.Regex.Replace(summary, "</a>", String.Empty);
                }

                FeedSubject.Add(subject);
                FeedSummary.Add(summary);
            }

            rssFeed_RichTextBox.Text = FeedSubject.ElementAt(0) + Environment.NewLine + FeedSummary.ElementAt(0);
            rssFeed_RichTextBox.Select(rssFeed_RichTextBox.GetFirstCharIndexFromLine(1), rssFeed_RichTextBox.Lines.ElementAt(1).Length);
            rssFeed_RichTextBox.SelectionFont = new Font("Arial", 16, FontStyle.Regular);

            while (readyBrowsers != 3)
            {
                Application.DoEvents();
            }


            mobilengine_Timer.Start();
            safetyLadder_Timer.Start();
            rssFeed_Timer.Start();
            //refreshDashboard_Timer.Start();

            splashScreen_Form.Close();
            this.Show();
        }

        private void RefreshDashboard_Timer_Tick(object sender, EventArgs e)
        {
            this.Refresh();
        }

        private void RssFeed_Timer_Tick(object sender, EventArgs e)
        {
            rssFeed_RichTextBox.Select(rssFeed_RichTextBox.GetFirstCharIndexFromLine(0), rssFeed_RichTextBox.Lines.ElementAt(0).Length);
            rssFeed_RichTextBox.SelectedText = FeedSubject.ElementAt(rssFeedCounter);
            rssFeed_RichTextBox.Select(rssFeed_RichTextBox.GetFirstCharIndexFromLine(1), rssFeed_RichTextBox.Lines.ElementAt(1).Length);
            rssFeed_RichTextBox.SelectedText = FeedSummary.ElementAt(rssFeedCounter);
            rssFeedCounter++;
            if (rssFeedCounter == (FeedSubject.Count - 1))
                rssFeedCounter = 0;
        }

        private void SafetyLadder_Timer_Tick(object sender, EventArgs e)
        {
            ((Timer)sender).Interval = 50;
            safetyLadder_Panel.AutoScrollPosition = new Point(0, safetyLadderScrollPos);
            Application.DoEvents();
            safetyLadderScrollPos++;
            if (safetyLadderScrollPos == safetyLadder_Panel.VerticalScroll.Maximum)
            {
                safetyLadderScrollPos = 0;
                safetyLadder_Panel.AutoScrollPosition = new Point(0, safetyLadderScrollPos);
                ((Timer)sender).Interval = 5000;
            }
        }

        private void Mobilengine_Timer_Tick(object sender, EventArgs e)
        {
            if (mobEngReset == false)
            {
                ((Timer)sender).Interval = 50;
                mobilengineContent.Parent.ScrollTop++;
                int mobEngOffset = mobilengineContent.Parent.ScrollRectangle.Height - mobilengineContent.Parent.ClientRectangle.Height;
                if (mobilengineContent.Parent.ScrollTop == mobEngOffset)
                {
                    ((Timer)sender).Interval = 5000; //pause at end of page and start of page
                    mobEngReset = true;
                }
            }
            else
            {
                mobilengineContent.Parent.ScrollTop = 0;
                mobEngReset = false;
            }

        }

        private void DateTimer_Tick(object sender, EventArgs e)
        {
            currentDateTime_RichTextBox.Text = DateTime.Now.ToShortTimeString() + Environment.NewLine + DateTime.Now.DayOfWeek + ", " + DateTime.Now.ToLongDateString();
            currentDateTime_RichTextBox.Select(currentDateTime_RichTextBox.GetFirstCharIndexFromLine(1), currentDateTime_RichTextBox.Lines.ElementAt(1).Length);
            currentDateTime_RichTextBox.SelectionFont = new Font("Arial", 22, FontStyle.Bold);
        }

        private void googleWeather_WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            googleWeather_WebBrowser.DocumentCompleted -= googleWeather_WebBrowser_DocumentCompleted;

            do //stall to load google's default cookies
            {
                Application.DoEvents();
            } while (googleWeather_WebBrowser.ReadyState != WebBrowserReadyState.Complete);

            string googleCookies = googleWeather_WebBrowser.Document.Cookie;
            string[] googleCookiesArray = System.Text.RegularExpressions.Regex.Split(googleCookies, "; ");
            foreach (string cookie in googleCookiesArray) //delete each cookie in the array
            {
                googleWeather_WebBrowser.Document.InvokeScript("eval", new object[] { "document.cookie = \"" + cookie + "; expires = Thu, 01 Jan 1970 00:00:00 UTC; domain=www.google.co.uk; path =/;\";" });
                googleWeather_WebBrowser.Document.InvokeScript("eval", new object[] { "document.cookie = \"" + cookie + "; expires = Thu, 01 Jan 1970 00:00:00 UTC; domain=.google.co.uk; path =/;\";" });

            }
            googleWeather_WebBrowser.Document.InvokeScript("eval", new object[] { "document.cookie = \"CONSENT=YES+GB.en-GB+V9\";" }); //add cookie consenting to google's terms of service
            googleWeather_WebBrowser.Refresh(); //refresh page, 

            do
            {
                Application.DoEvents();
            } while (googleWeather_WebBrowser.ReadyState != WebBrowserReadyState.Complete);

            bool googleWeatherContinue = false;
            while (!googleWeatherContinue)
            {
                try
                {
                    HtmlElement googleWeatherElement = ElementsByClass(googleWeather_WebBrowser.Document, "vk_c card-section").First();
                    googleWeather_WebBrowser.Document.Body.Style = "zoom:80%;";
                    googleWeatherElement.ScrollIntoView(true);
                    googleWeatherContinue = true;
                }
                catch
                {
                    Application.DoEvents();
                }
            }
            readyBrowsers++;
            splashScreen_Label.Text += " Loaded Weather...";
        }

        public void SaveEditor(string editor)
        {
            System.IO.File.WriteAllText(editorDocPath, editor);
            editorDoc.InvokeScript("eval", new object[]{ "location.reload()" });                              //eval runs this script without a matching function being present in the target document
            editorDocText = editor;
        }

        public void ReturnEditorSize(int width, int height)
        {
            mcgeeEditor_WebBrowser.Size = new Size(width, height);
            mcgeeEditor_Table.Size = new Size(mcgeeEditor_WebBrowser.Size.Width, mcgeeEditor_WebBrowser.Size.Height + 30);
        }

        private void twitter_WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            twitter_WebBrowser.DocumentCompleted -= twitter_WebBrowser_DocumentCompleted;
            bool twitterContinue = false;
            while (!twitterContinue)
            {
                try
                {
                    twitter_WebBrowser.Document.Body.Style = "zoom:85%;";
                    HtmlElement twitterHeader = ElementsByClass(twitter_WebBrowser.Document, "topbar js-topbar").First();
                    twitterHeader.Style = "visibility:hidden;";
                    HtmlElement twitterFeed = ElementsByClass(twitter_WebBrowser.Document, "ProfileHeading").First();
                    twitterFeed.ScrollIntoView(true);
                    twitterContinue = true;
                }
                catch
                {
                    Application.DoEvents();
                }
            }
            readyBrowsers++;
            splashScreen_Label.Text += " Loaded Twitter...";
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                if (MessageBox.Show("Exit the McGee Dashboard?", "McGee Dashboard", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    Application.Exit();
                }

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void WebBrowser_DocumentClicked(object sender, HtmlElementEventArgs e)
        {
            if (editorActive)
                return;

            editorActive = true;
            editorDoc = (HtmlDocument)sender;
            editorDoc.Click -= WebBrowser_DocumentClicked;
            editorDocPath = editorDoc.Url.LocalPath;
            editorDocText = System.IO.File.ReadAllText(editorDocPath);
            mcgeeEditor_WebBrowser.Document.InvokeScript("SetEditor", new Object[] { editorDoc.DomDocument });
            mcgeeEditor_WebBrowser.Size = editorDoc.Body.ClientRectangle.Size;
            mcgeeEditor_WebBrowser.Document.InvokeScript("SetEditorSize", new Object[] { mcgeeEditor_WebBrowser.Size.Width, mcgeeEditor_WebBrowser.Size.Height });
            Controls.Add(this.mcgeeEditor_Table);

            do
            {
                Application.DoEvents();
            } while (mcgeeEditor_WebBrowser.ReadyState != WebBrowserReadyState.Complete);

            mcgeeEditor_Table.Visible = true;
            mcgeeEditor_Table.BringToFront();
        }

        private void mcgeeEditorCloseBtn_PictureBox_Click(object sender, EventArgs e)
        {
            string currentEditorText = mcgeeEditor_WebBrowser.Document.InvokeScript("GetEditor").ToString();
            //MessageBox.Show("Current Editor Text" + Environment.NewLine + currentEditorText + "Original Document Text" + Environment.NewLine + editorDocText);
            if (currentEditorText != editorDocText)
            {
                if (MessageBox.Show("Exit McGee Dashboard Editor without saving changes?", "McGee Dashboard", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    CloseEditor();

            }

            else
                CloseEditor();
        }

        private void CloseEditor()
        {
            this.Controls.Remove(mcgeeEditor_Table);
            editorActive = false;
            editorEnlarged = false;
            editorDoc.Click += WebBrowser_DocumentClicked;
        }

        private void mcgeeEditorResizeBtn_PictureBox_Click(object sender, EventArgs e)
        {

            mcgeeEditor_WebBrowser.Size = editorDoc.Body.ClientRectangle.Size;

            if (!editorEnlarged)
            {
                mcgeeEditor_WebBrowser.Document.InvokeScript("SetEditorSize", new Object[] { mcgeeEditor_WebBrowser.Size.Width /*same width*/ , 750 /*new height*/});
                editorEnlarged = true;
            }
            else
            {
                mcgeeEditor_WebBrowser.Document.InvokeScript("SetEditorSize", new Object[] { mcgeeEditor_WebBrowser.Size.Width, mcgeeEditor_WebBrowser.Size.Height });
                editorEnlarged = false;
            }

        }

        private void mcgeeEditorResizeBtn_PictureBox_Rollover(object sender, EventArgs e)
        {
            this.mcgeeEditorResizeBtn_PictureBox.Image = global::NSMcGeeDashboard.Properties.Resources.Resize_Button_Rollover;
            Cursor = Cursors.Hand;
        }

        private void mcgeeEditorResizeBtn_PictureBox_Rolloff(object sender, EventArgs e)
        {
            this.mcgeeEditorResizeBtn_PictureBox.Image = global::NSMcGeeDashboard.Properties.Resources.Resize_Button;
            Cursor = DefaultCursor;
        }

        private void mcgeeEditorCloseBtn_PictureBox_Rollover(object sender, EventArgs e)
        {
            this.mcgeeEditorCloseBtn_PictureBox.Image = global::NSMcGeeDashboard.Properties.Resources.Close_Button_Rollover;
            Cursor = Cursors.Hand;
        }

        private void mcgeeEditorCloseBtn_PictureBox_Rolloff(object sender, EventArgs e)
        {
            this.mcgeeEditorCloseBtn_PictureBox.Image = global::NSMcGeeDashboard.Properties.Resources.Close_Button;
            Cursor = DefaultCursor;
        }

        private void mobilengine_WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            //Original mobilengine Url... page is now down?
            //https://app.periscopedata.com/shared/1edefb9a-400b-4eec-8b10-2864de5e2d6f

            mobilengine_WebBrowser.DocumentCompleted -= mobilengine_WebBrowser_DocumentCompleted;
            bool mobilengineContinue = false;
            while (!mobilengineContinue)
            {
                try
                {
                    // For original Url...
                    /*
                    mobilengine_WebBrowser.Document.Body.Style = "zoom:115%;";
                    mobilengineHeader = ElementsByClass(mobilengine_WebBrowser.Document, "app-header dashboard-header tooltip-cover-heading").First();
                    mobilengineHeader.Style = "visibility:hidden;";
                    */

                    mobilengineContent = ElementsByClass(mobilengine_WebBrowser.Document, "main-content").First();
                    mobilengineContent.Style = "overflow:hidden; top:0;";
                    mobilengine_WebBrowser.Visible = true;
                    mobilengineContinue = true;
                }
                catch
                {
                    Application.DoEvents();
                }
            }

            readyBrowsers++;
            splashScreen_Label.Text += " Loaded Mobilengine...";
        }
    }
}
