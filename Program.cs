using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Group;
using System.Diagnostics;

namespace LyncFellow
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            using (Mutex mutexApplication = new Mutex(false, "LyncFellowApplication"))
            {
                if (!mutexApplication.WaitOne(0, false))
                {
                    MessageBox.Show(Application.ProductName + " is already running!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                ApplicationContext applicationContext = new ApplicationContext();
                Application.Run(applicationContext);
            }
        }
    }

    class ApplicationContext : System.Windows.Forms.ApplicationContext
    {

        System.ComponentModel.IContainer _components;
        NotifyIcon _notifyIcon;
        SettingsForm _settingsForm;

        Buddies _buddies;
        LyncClient _lyncClient;
        System.Windows.Forms.Timer _housekeepingTimer;
        DateTime _lyncEventsUpdated;
        public bool thisDebug = false;


        string[] ColorListForMenu = { "Off", "Black", "Red", "Lilac", "Pink", "Green", "LightGreen", "Cyan", "LightBlue" };

        public ApplicationContext()
        {

            if (Properties.Settings.Default.CallUpgrade)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.CallUpgrade = false;
                Properties.Settings.Default.Save();
            }

            _components = new System.ComponentModel.Container();
            _notifyIcon = new NotifyIcon(_components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Visible = true
            };


            Trace.WriteLine("Available color: " + Properties.Settings.Default.AvailableColor);

            //handle doubleclicks on the icon:
            _notifyIcon.DoubleClick += TrayIcon_DoubleClick;

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Settings", null, new EventHandler(MenuSettingsItem_Click)));
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Exit", null, new EventHandler(MenuExitItem_Click)));
            //_notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Dance Now", null, new EventHandler(MenuDance_Click)));
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("HeartBeat Now", null, new EventHandler(MenuHeartBeat_Click)));
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("FlapWings Now", null, new EventHandler(FlapWings_Click)));
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Rainbox It", null, new EventHandler(MenuRainbox_Click)));


            //get names from group



            //go through each group and pull out as much info as we want..

            /*
                foreach (Group _Group in LyncClient.GetClient().ContactManager.Groups)
                {
                    Trace.WriteLine("_Group = " + _Group);
                    GetGroupContacts(_Group);
   
                }
            */

            /*
             SIMPLE DIALOG BOX
 
            DialogResult result1 = MessageBox.Show("Is Dot Net Perls awesome?", "Important Question", MessageBoxButtons.YesNo);
            //Trace.WriteLine("TRACE ME: State='{0}", result1);
 
             if (result1.ToString()=="Yes")
             {
                 MessageBox.Show("YES");
                 Trace.WriteLine("TRACE YES");

             }
             else {
                 Trace.WriteLine("TRACE NO");
             }
                Environment.Exit(-1);
            */


            //Create the menu node and get Webteam and add to Menu bar
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("WebTeam", null, new EventHandler(ToolStripMenuItem5_DoubleClick)) { Image = Properties.Resources.webTeamIcon.ToBitmap() });
            getGroup("webteam", null, 5, "webTeam");

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Sales - OSS", null, new EventHandler(ToolStripMenuItem6_DoubleClick)) { Image = Properties.Resources.salesTeamOss.ToBitmap() });
            getGroup("salesTeamOSS", null, 6, "Sales - OSS");

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Sales - Pro Audio", null, new EventHandler(ToolStripMenuItem7_DoubleClick)) { Image = Properties.Resources.salesTeamTmp.ToBitmap() });
            getGroup("salesProAudio", null, 7, "Sales - Pro Audio");

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("All Employees", null, new EventHandler(ToolStripMenuItem8_DoubleClick)) { Image = Properties.Resources.salesTeamTmp.ToBitmap() });
            getGroup("salesAllEmployees", null, 8, "All Employees");


            /* TEST COLORS SUB MENU, in an Array  */
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripMenuItem("Test Colors", null, new EventHandler(TestColors_Click)));
            foreach (string eachColor in ColorListForMenu)
            {

                ToolStripMenuItem item = new ToolStripMenuItem(eachColor) { Image = Properties.Resources.webTeamIcon.ToBitmap() };
                (_notifyIcon.ContextMenuStrip.Items[9] as ToolStripMenuItem).DropDownItems.Add(item);
                item.Click += new EventHandler(TestColors_Click);

            }

            //INITILIZE BUDDY and HOUSEKEEPING
            _buddies = new Buddies();
            _buddies.Rainbow(500);  //woo

            _buddies.users.Add("sip:pk@musicpeopleinc.com");
            _buddies.users.Add("sip:ajr@musicpeopleinc.com");
            _buddies.users_group.Add("sip:ctc@musicpeopleinc.com");

            _housekeepingTimer = new System.Windows.Forms.Timer();
            _housekeepingTimer.Interval = 15000;
            _housekeepingTimer.Tick += HousekeepingTimer_Tick;
            HousekeepingTimer_Tick(null, null);     // tick anyway enables timer when finished


        }

        private void ToolStripMenuItem5_DoubleClick(Object sender, EventArgs e)
        {
            //webteam
            foreach (Group _Group in LyncClient.GetClient().ContactManager.Groups)
            {
                //Trace.WriteLine("_Group = " + _Group);
                if (_Group.Name == "webteam")
                {
                    GetGroupContacts(_Group);

                    _buddies.users_group.Add("sip:chg@musicpeopleinc.com");
                    _buddies.users_group.Add("sip:ajr@musicpeopleinc.com");
                    _buddies.users_group.Add("sip:ctc@musicpeopleinc.com");

                    //show Display Box With Custom List
                    var userWindowDisplay = new LyncFellow.Window1(_buddies.users_group);
                    userWindowDisplay.Title = "TMP - Webteam Group";
                    userWindowDisplay.Show();
                }
            }
        }


        private void ToolStripMenuItem6_DoubleClick(Object sender, EventArgs e)
        {
            //webteam
            foreach (Group _Group in LyncClient.GetClient().ContactManager.Groups)
            {
                //Trace.WriteLine("_Group = " + _Group);
                if (_Group.Name == "Sales - OSS")
                {
                    GetGroupContacts(_Group);
                    //show Display Box With Custom List
                    var userWindowDisplay = new LyncFellow.Window1(_buddies.users_group);
                    userWindowDisplay.Title = "TMP - Sales - OSS Group";
                    userWindowDisplay.Show();
                }
            }
        }
        private void ToolStripMenuItem7_DoubleClick(Object sender, EventArgs e)
        {
            //webteam
            foreach (Group _Group in LyncClient.GetClient().ContactManager.Groups)
            {
                //Trace.WriteLine("_Group = " + _Group);
                if (_Group.Name == "Sales - Pro Audio")
                {
                    GetGroupContacts(_Group);
                    //show Display Box With Custom List
                    var userWindowDisplay = new LyncFellow.Window1(_buddies.users_group);
                    userWindowDisplay.Title = "TMP - Sales - Pro Audio";
                    userWindowDisplay.Show();
                }
            }
        }
        private void ToolStripMenuItem8_DoubleClick(Object sender, EventArgs e)
        {
            //webteam
            foreach (Group _Group in LyncClient.GetClient().ContactManager.Groups)
            {
                //Trace.WriteLine("_Group = " + _Group);
                if (_Group.Name == "All Employees")
                {
                    GetGroupContacts(_Group);
                    //show Display Box With Custom List
                    var userWindowDisplay = new LyncFellow.Window1(_buddies.users_group);
                    userWindowDisplay.Title = "TMP - All Employees";
                    userWindowDisplay.Show();
                }
            }
        }

        //getGroup("salesTeamOSS", null, 6, "Sales - OSS");
        private void getGroup(string passedGroupTmp, Group passedGroup, int menuCount, string groupName)
        {

            (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Clear(); //clear all the menu iteams and remake

            if (LyncClient.GetClient().ContactManager.Groups.TryGetGroup(groupName, out passedGroup))
            {
                foreach (Contact salesPerson in passedGroup)
                {
                    ContactModel webTeamMember = new ContactModel(salesPerson);
                    object activityString = salesPerson.GetContactInformation(ContactInformationType.Activity);
                    object DisplayName = salesPerson.GetContactInformation(ContactInformationType.DisplayName);

                    //Trace.WriteLine(string.Format("activityString.ToString() = \"{0}\"", activityString.ToString()));
                    //Trace.WriteLine(string.Format("DisplayName = \"{0}\"", DisplayName.ToString()));

                    if (activityString.ToString() == "Available" || activityString.ToString() == "Available for ACD - Available")
                    {
                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.available.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "Offline")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.offline.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "In a meeting")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.meeting.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "Away")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.inactive.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "Inactive")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.inactive.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "In a call")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.In_a_call.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "Busy")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.busy.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }
                    else if (activityString.ToString() == "Off work")
                    {

                        ToolStripMenuItem item = new ToolStripMenuItem("" + webTeamMember) { Image = Properties.Resources.offline.ToBitmap() };
                        (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add(item);
                        item.Click += new EventHandler(TryUser_Click);
                    }

                    //Trace.WriteLine(string.Format("TRACE ME: activityString=\"{0}\"", activityString.ToString()));
                }

                //add chris Gobel
                if (groupName == "webTeam")
                {
                    (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add("Chris H. Goebel - TMP", null, new EventHandler(TryUser_Click));
                    (_notifyIcon.ContextMenuStrip.Items[menuCount] as ToolStripMenuItem).DropDownItems.Add("Aaron J. Ryan - TMP", null, new EventHandler(TryUser_Click));
                }

            }

        }//getGroup

        // Gets and processes each group of contacts for extended inforamtion, Currently just here
        private void GetGroupContacts(Group group)
        {
            //Trace.WriteLine("GROUP = " + group);
            _buddies.users_group.Clear();

            // Iterate on the contacts in the passed group.
            foreach (Contact _Contact in group)
            {

                // Get contact information from the contact.
                string uri = _Contact.Uri;
                object activityString = _Contact.GetContactInformation(ContactInformationType.Activity);
                object DisplayName = _Contact.GetContactInformation(ContactInformationType.DisplayName);
                object Desc = _Contact.GetContactInformation(ContactInformationType.Description);
                object Availability = _Contact.GetContactInformation(ContactInformationType.Availability);
                //object photoTest = _Contact.GetContactInformation(ContactInformationType.Photo);

                //var DisplayName = ContactInformationType.DisplayName;
                _buddies.users_group.Add(_Contact.Uri);
                //Trace.WriteLine("URI=" + uri + " activityString=" + activityString + "DisplayName = " + DisplayName + " photoTest" + "Desc = " + Desc + " Availability=" + Availability);
                //Trace.WriteLine(_Contact.GetContactInformation(ContactInformationType.Photo));

            }

        }

        // Gets SIP for display name

        private void HousekeepingTimer_Tick(object sender, EventArgs e)
        {
            _housekeepingTimer.Enabled = false;

            _buddies.RefreshList();

            if (_lyncClient != null && _lyncClient.State == ClientState.Invalid)
            {
                Trace.WriteLine("LyncFellow: _lyncClient != null && _lyncClient.State == ClientState.Invalid");
                ReleaseLyncClient();
            }
            if (_lyncClient == null)
            {
                try
                {
                    _lyncClient = LyncClient.GetClient();
                }
                catch { }
                if (_lyncClient != null)
                {
                    if (_lyncClient.State != ClientState.Invalid)
                    {
                        if (thisDebug)
                        {
                            Trace.WriteLine("LYNC STATE CHANGED");
                        }
                        _lyncClient.StateChanged += new EventHandler<ClientStateChangedEventArgs>(LyncClient_StateChanged);

                        Trace.WriteLine("SENDING #1 -> ConversationManager_ConversationAdded");

                        _lyncClient.ConversationManager.ConversationAdded += new EventHandler<ConversationManagerEventArgs>(ConversationManager_ConversationAdded);
                        HandleSelfContactEventConnection(_lyncClient.State);
                    }
                    else
                    {
                        ReleaseLyncClient();
                    }
                }
            }
            if (_lyncClient != null && DateTime.Now.Subtract(_lyncEventsUpdated).TotalMinutes > 60)
            {
                if (thisDebug)
                    Trace.WriteLine("LyncFellow: _lyncClient != null && DateTime.Now.Subtract(_lyncEventsUpdated).TotalMinutes > 60");

                HandleSelfContactEventConnection(_lyncClient.State);
            }


            if (_buddies.currentUser == "self")
            {
                UpdateBuddiesColorBySelfAvailability();
            }
            else
            {
                UpdateBuddiesColorByUserAvailability(_buddies.currentUser);
            }


            bool deviceNotFunctioning = _buddies.LastWin32Error == 0x1f;    //ERROR_GEN_FAILURE
            bool lyncValid = IsLyncConnectionValid();
            if (_buddies.Count > 0 && !deviceNotFunctioning && lyncValid)
            {
                _notifyIcon.Text = Application.ProductName;
                _notifyIcon.Icon = Properties.Resources.LyncFellow;
            }
            else
            {
                _notifyIcon.Text = Application.ProductName;
                _notifyIcon.Icon = Properties.Resources.LyncFellowInfo;
                if (_buddies.Count == 0)
                {
                    _notifyIcon.Text += "\r* No USB buddy found.";
                }
                else if (deviceNotFunctioning)
                {
                    //_notifyIcon.Text += "\r* Please reconnect USB buddy.";
                    _notifyIcon.Icon = Properties.Resources.LyncFellow;

                }
                if (!lyncValid)
                {
                    _notifyIcon.Text += "\r* No Lync connection.";
                }
            }

            _housekeepingTimer.Enabled = true;

            //update menu node items
            getGroup("webteam", null, 5, "webTeam");
            getGroup("salesTeamOSS", null, 6, "Sales - OSS");
            getGroup("salesProAudio", null, 7, "Sales - Pro Audio");
            getGroup("salesAllEmployees", null, 8, "All Employees");

        }

        private bool IsLyncConnectionValid()
        {
            return _lyncClient != null && _lyncClient.State == ClientState.SignedIn && _lyncClient.Self.Contact != null;
        }

        private void ReleaseLyncClient()
        {
            _lyncClient = null;
            GC.Collect();
        }

        private void HandleSelfContactEventConnection(ClientState State)
        {
            if (State == ClientState.SignedIn || State == ClientState.SigningOut)
            {
                _lyncClient.Self.Contact.ContactInformationChanged -= new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);
                //Trace.WriteLine("LyncFellow: _lyncClient.Self.Contact.ContactInformationChanged -= SelfContact_ContactInformationChanged");

                if (State == ClientState.SignedIn)
                {
                    _lyncClient.Self.Contact.ContactInformationChanged += new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);
                    _lyncEventsUpdated = DateTime.Now;
                    //Trace.WriteLine("LyncFellow: _lyncClient.Self.Contact.ContactInformationChanged += SelfContact_ContactInformationChanged");
                }
            }
        }

        private void LyncClient_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            Trace.WriteLine(string.Format("LyncFellow: LyncClient_StateChanged e.OldState={0}, e.NewState={1}, e.StatusCode=0x{2:x}", e.OldState, e.NewState, e.StatusCode));
            HandleSelfContactEventConnection(e.NewState);
            UpdateBuddiesColorBySelfAvailability();
        }

        private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            foreach (ContactInformationType changedInformationType in e.ChangedContactInformation)
            {
                if (changedInformationType == ContactInformationType.Availability)
                {
                    UpdateBuddiesColorBySelfAvailability();
                }
            }
        }

        private void ConversationManager_ConversationAdded(object sender, ConversationManagerEventArgs e)
        {

            Trace.WriteLine("TEST ConversationManager_ConversationAdded");
            Trace.WriteLine("SENDER -> " + sender);

            bool IncomingCall = false;

            if (e.Conversation.Participants.Count >= 2)
            {
                var InitiatorTest = e.Conversation.Participants[1].Contact;
                Console.WriteLine("IM STARTED InitiatorTestUri = " + InitiatorTest);
                _buddies.FlapWings(5000);
                _buddies.Rainbow(5000);
                _buddies.Heartbeat(10000);

            }



            /*
             * 
             *             foreach (var Modality in e.Conversation.Modalities)
            {    
                Console.WriteLine(Modality.Value.ToString());
            }
             * 
                IDictionary<InstantMessageContentType, string> messageFormatProperty = e.Contents;
                if (messageFormatProperty.ContainsKey(InstantMessageContentType.PlainText))
                {
                    string outVal = string.Empty;
                    MessageBox.Show("New message: " + outVal);
                }

             */

            foreach (var Modality in e.Conversation.Modalities)
            {
                if (Modality.Value != null && Modality.Value.State == ModalityState.Notified)
                {
                    IncomingCall = true;
                }
            }


            if (IncomingCall)
            {
                //IncomingCall Detected

                if (Properties.Settings.Default.DanceOnIncomingCall)
                {
                    _buddies.Dance(3500);
                    _buddies.FlapWings(3500);
                }

                if (e.Conversation.Participants.Count >= 2)
                {
                    var Initiator = e.Conversation.Participants[1].Contact;

                    //MessageBox.Show("Someone is calling you from Uri = " + Initiator.Uri);

                    ShowBalloonText("Call from Uri: " + Initiator.Uri);

                    //Trace.WriteLine(string.Format("TOP - LyncFellow: Initiator.Uri=\"{0}\"", Initiator.Uri));
                    // magic heartbeat for incoming conversations from G&K ;-)
                    if (Initiator.Uri.Contains("+18605153636"))
                    {
                        _buddies.Heartbeat(10000);
                        _buddies.FlapWings(5000);
                        _buddies.FreakOut(1000);
                    }
                    else
                    {
                        _buddies.Rainbow(5000);
                        _buddies.Heartbeat(10000);
                    }
                }
            }
            else
            {

                if (e.Conversation.Participants.Count >= 2)
                {
                    var Initiator = e.Conversation.Participants[1].Contact;

                    // MessageBox.Show("Someone is calling you from Uri = " + Initiator.Uri);
                    ShowBalloonText("Starting a call with: \r" + Initiator.Uri.Replace("sip:", ""));

                }


                //Incoming Meeting Detected, FREAK OUT
                _buddies.Rainbow(10000);
                _buddies.Heartbeat(10000);
                _buddies.FreakOut(1000);

                //MessageBox.Show("Looks like a conference call is about to happen!");
                //ShowBalloonText("Looks like a conference call is about to happen!");

            }
        }

        private void UpdateBuddiesColorBySelfAvailability()
        {

            UpdateBuddiesColorByUserAvailability(_lyncClient.Self.Contact.GetContactInformation(ContactInformationType.DisplayName).ToString());

        }

        private void ShowBalloonText(string text)
        {
            //MessageBox.Show("A call from: " + text);

            //show Balloon Icon
            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
            _notifyIcon.BalloonTipTitle = "Lync";

            _notifyIcon.BalloonTipText = text;


            //Here you can do stuff if the tray icon is doubleclicked
            _notifyIcon.ShowBalloonTip(10000);

        }

        private void MenuSettingsItem_Click(object sender, EventArgs e)
        {

            if (_settingsForm == null)
            {
                _settingsForm = new SettingsForm();
                _settingsForm.Closed += settingsForm_Closed;
                _settingsForm.Show();
            }
            else { _settingsForm.Activate(); }
        }

        void settingsForm_Closed(object sender, EventArgs e)
        {
            _settingsForm = null;

            UpdateBuddiesColorBySelfAvailability();
        }

        private void MenuExitItem_Click(object sender, EventArgs e)
        {
            ExitThread();
        }

        private void MenuHeartBeat_Click(object sender, EventArgs e)
        {
            _buddies.Heartbeat(10000);
        }

        private void MenuRainbox_Click(object sender, EventArgs e)
        {
            _buddies.Rainbow(5000);
        }

        private void MenuDance_Click(object sender, EventArgs e)
        {
            _buddies.Rainbow(5000);
            _buddies.Heartbeat(5000);
            _buddies.Dance(5000);
            _buddies.FlapWings(5000);
        }

        private void FlapWings_Click(object sender, EventArgs e)
        {
            _buddies.FlapWings(3500);
            _buddies.Dance(1250);

        }

        private void TryUser_Click(object sender, EventArgs e)
        {

            //get TeamMember from click
            var teamMember = ((System.Windows.Forms.ToolStripDropDownItem)(sender));

            //uncheck all item, in all groups
            for (int i = 5; i < 9; i++)
            {
                foreach (ToolStripMenuItem item in (_notifyIcon.ContextMenuStrip.Items[i] as ToolStripMenuItem).DropDownItems)
                {
                    item.Checked = false;
                }
            }

            ((ToolStripMenuItem)teamMember).Checked = true;

            _buddies.currentUser = teamMember.ToString().Replace("WebTeam Member: ", "");

            UpdateBuddiesColorByUserAvailability(_buddies.currentUser);

        }

        private void TestColors_Click(object sender, EventArgs e)
        {

            Trace.WriteLine("TestColors_Click");


            var menuItem = ((System.Windows.Forms.ToolStripDropDownItem)(sender));

            Trace.WriteLine("HERE BITCH = " + menuItem);
            if (menuItem.ToString().Contains("Black"))
            {
                _buddies.Color = iBuddy.Color.Black;
                Trace.WriteLine("MADE IT Black - 7");
            }

            if (menuItem.ToString().Contains("Red"))
            {
                _buddies.Color = iBuddy.Color.Red;
                Trace.WriteLine("MADE IT RED - 6");
            }

            if (menuItem.ToString().Contains("Green"))
            {
                _buddies.Color = iBuddy.Color.Green;
                Trace.WriteLine("MADE IT Green - 5");
            }

            if (menuItem.ToString().Contains("Pink"))
            {
                _buddies.Color = iBuddy.Color.Pink;
                Trace.WriteLine("MADE IT Pink - 4");
            }

            if (menuItem.ToString().Contains("Blue"))
            {
                _buddies.Color = iBuddy.Color.LightBlue;
                Trace.WriteLine("MADE IT LightBlue - 3 ");
            }

            if (menuItem.ToString().Contains("Lilac"))
            {
                _buddies.Color = iBuddy.Color.Lilac;
                Trace.WriteLine("MADE IT Lila - 2");
            }

            if (menuItem.ToString().Contains("Cyan"))
            {
                _buddies.Color = iBuddy.Color.Cyan;
                Trace.WriteLine("MADE IT Cyan - 1 ");
            }

            if (menuItem.ToString().Contains("LightBlue"))
            {
                _buddies.Color = iBuddy.Color.LightBlue;
                Trace.WriteLine("MADE IT LightBlue - 0");
            }
        }

        private void UpdateBuddiesColorByUserAvailability(string user)
        {

            //Trace.WriteLine(string.Format("TRACE ME: GetSipForDisplayName=\"{0}\"", GetSipForDisplayName(_buddies.currentUser)));
            // _buddies.users.Add(GetSipForDisplayName(_buddies.currentUser));

            _notifyIcon.Text = "Current user: " + user;

            if (thisDebug)
            {
                Trace.WriteLine(string.Format("START: user=\"{0}\"", user));
            }

            ContactAvailability Availability = ContactAvailability.None;
            string Activity = "";

            var searchLyncForPassedUser = LyncClient.GetClient().ContactManager.BeginSearch(
                user,
                (ar) =>
                {
                    SearchResults searchResults = LyncClient.GetClient().ContactManager.EndSearch(ar);
                    if (searchResults.Contacts.Count > 0)
                    {
                        //Console.WriteLine(searchResults.Contacts.Count.ToString() +" found");

                        foreach (Contact contact in searchResults.Contacts)
                        {
                            //found the user
                            if (thisDebug)
                            {
                                Console.WriteLine(contact.GetContactInformation(ContactInformationType.DisplayName).ToString());
                            }

                            //pull extended information from user
                            Availability = (ContactAvailability)contact.GetContactInformation(ContactInformationType.Availability);
                            Activity = (string)contact.GetContactInformation(ContactInformationType.ActivityId);
                            _buddies.currentUserUri = (string)contact.Uri;

                            _buddies.ActivityString = contact.GetContactInformation(ContactInformationType.Activity);
                            _buddies.DisplayName = contact.GetContactInformation(ContactInformationType.DisplayName);
                            _buddies.Desc = contact.GetContactInformation(ContactInformationType.Description);
                            _buddies.Availability = contact.GetContactInformation(ContactInformationType.Availability);
                            _buddies.PersonalNote = contact.GetContactInformation(ContactInformationType.PersonalNote);

                            _notifyIcon.Text = "Current user: " + user + " -  " + Activity;

                            _buddies.currentUserName = user;
                            _buddies.currentUserActivity = Activity;

                            _notifyIcon.Icon = Properties.Resources.LyncFellowInfo;


                            if (!_buddies.users.Contains(contact.Uri))
                            {
                                Trace.WriteLine("ADDING TO LIST -->_buddies.currentUserUri" + _buddies.currentUserUri);

                                _buddies.users.Add(_buddies.currentUserUri);
                            }

                            /*

                            if (_buddies.currentUser.Contains("Peter Kujawa"))
                            {
                                _buddies.currentUser = "self";
                                _notifyIcon.Icon = Properties.Resources.LyncFellow;
                            }
                            else
                            {
                                if (_buddies.currentUserUri != "self")
                                {
                                    Trace.WriteLine("ADDING TO LIST -->_buddies.currentUserUri" + _buddies.currentUserUri);

                                    _buddies.users.Add(_buddies.currentUserUri);

                                }
                            }
                             */

                            if (thisDebug)
                            {
                                Console.WriteLine("user = " + user);
                                Console.WriteLine("Availability = " + Availability);
                                Console.WriteLine("Activity = " + Activity);
                                Console.WriteLine("");
                            }

                            bool RedOnBusy = Properties.Settings.Default.RedOnDndCallBusy == ContactAvailability.Busy;
                            bool RedOnCall = RedOnBusy || Properties.Settings.Default.RedOnDndCallBusy == ContactAvailability.None;
                            bool InACall = Activity == "on-the-phone" || Activity == "in-a-conference";

                            bool Dnd = Availability == ContactAvailability.DoNotDisturb;
                            bool Busy = Availability == ContactAvailability.Busy;
                            bool BusyIdle = Availability == ContactAvailability.BusyIdle;
                            bool Free = Availability == ContactAvailability.Free;
                            bool FreeIdle = Availability == ContactAvailability.FreeIdle;
                            bool Away = Availability == ContactAvailability.Away;
                            bool TemporarilyAway = Availability == ContactAvailability.TemporarilyAway;

                            var colorOld = _buddies.Color;
                            var colorNew = iBuddy.Color.Off;

                            if (Dnd || (RedOnCall && InACall) || (RedOnBusy && Busy))
                            {
                                colorNew = iBuddy.Color.Red;
                            }
                            else if (Away)
                            {
                                colorNew = iBuddy.Color.Lilac;
                            }
                            else if (Busy)
                            {
                                colorNew = iBuddy.Color.Pink;
                            }
                            else if (BusyIdle)
                            {
                                colorNew = iBuddy.Color.Cyan;
                            }
                            else if (Free)
                            {
                                if (Properties.Settings.Default.AvailableColor == "Green")
                                {
                                    colorNew = iBuddy.Color.Green;
                                }

                                if (Properties.Settings.Default.AvailableColor == "Lilac")
                                {
                                    colorNew = iBuddy.Color.Lilac;
                                }

                                if (Properties.Settings.Default.AvailableColor == "Cyan")
                                {
                                    colorNew = iBuddy.Color.Cyan;
                                }

                                if (Properties.Settings.Default.AvailableColor == "Pink")
                                {
                                    colorNew = iBuddy.Color.Pink;
                                }

                            }
                            else if (FreeIdle)
                            {
                                colorNew = iBuddy.Color.LightGreen;
                            }
                            else if (TemporarilyAway)
                            {
                                colorNew = iBuddy.Color.LightBlue;
                            }

                            //set the color
                            _buddies.Color = colorNew;

                            var colorNewString = _buddies.Color;

                            if (colorOld != colorNew)
                            {
                                //if color is changing, then rainbow it
                                _buddies.Rainbow(2000);
                            }

                            //if they were red and are now available then do it
                            if (colorOld.ToString() == "Red" & colorNewString.ToString() == "Free")
                            {
                                _buddies.Rainbow(2000);
                                _buddies.Heartbeat(2000);
                                _buddies.Dance(2000);
                                _buddies.FlapWings(2000);
                                Trace.WriteLine("NO MORE WAITING");
                            }


                            if (thisDebug)
                            {
                                Trace.WriteLine(string.Format("LyncFellow: UpdateBuddiesColorByUserAvailability"));
                            }


                            break;

                        }
                    }
                },
                null);

            //SearchForContacts(user);

            //MessageBox.Show("RING: " + user);

        }


        protected override void ExitThreadCore()
        {
            if (_settingsForm != null) { _settingsForm.Close(); }
            _settingsForm.Close();
            _notifyIcon.Visible = false;
            base.ExitThreadCore();

        }


        //TEST clicks
        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {

            //show Display Box
            var userWindowDisplay = new LyncFellow.Window1(_buddies.users);
            userWindowDisplay.Title = "TMP Lync - Custom List";
            Uri iconUri = new Uri("pack://Resources/worker.ico", UriKind.RelativeOrAbsolute);

            userWindowDisplay.Show();

            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
            _notifyIcon.BalloonTipTitle = "Current User: " + _buddies.currentUserName + "\r";

            _notifyIcon.BalloonTipText = "Activity: " + _buddies.currentUserActivity + " / " + _buddies.ActivityString + "\r";
            _notifyIcon.BalloonTipText += "DisplayName: " + _buddies.DisplayName + "\r";
            _notifyIcon.BalloonTipText += "Desc: " + _buddies.Desc + "\r";
            _notifyIcon.BalloonTipText += "Availability: " + _buddies.Availability + "\r";

            //Here you can do stuff if the tray icon is doubleclicked
            _notifyIcon.ShowBalloonTip(10000);


        }


    }

    //Internal class for Lync Contacts
    internal class ContactModel
    {
        private Contact thisContact;
        public Contact LyncContact
        {
            get
            {
                return thisContact;
            }
        }
        public override string ToString()
        {
            return thisContact.GetContactInformation(ContactInformationType.DisplayName).ToString();
        }
        public ContactModel(Contact aContact)
        {
            thisContact = aContact;
        }
    }

}
