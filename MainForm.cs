using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace GratisInc.Tools.FogBugz.WorkingOn
{
    public partial class MainForm : Form
    {
        private Case workingCase;
        private Dictionary<Int32, String> projects = new Dictionary<Int32, String>();        

        public MainForm()
        {
            InitializeComponent();
            tbManualCase.TextBox.LostFocus += new EventHandler(TextBox_LostFocus);

            // Position the form to the bottom right corner.
            Int32 x = SystemInformation.WorkingArea.Right - this.Width;
            Int32 y = SystemInformation.WorkingArea.Bottom - this.Height;
            this.SetDesktopLocation(x, y);

            // Load any saved settings into the form.
            tbServer.Text = Settings.Default.Server;
            tbUser.Text = Settings.Default.User;
            tbPassword.Text = Settings.Default.Password;

            // If the settings are all present, automatically initiate the logon process.
            if (tbServer.Text != "" && tbUser.Text != "" && tbPassword.Text != "")
            {
                Logon();
                UpdateProjects();
                UpdateWorkingCase();
                UpdateCases();
            }
        }
                
        #region [ FogBugz API Calls ]

        /// <summary>
        /// Initiates the logon process. If it is successful, the form values
        /// are saved to the application settings.
        /// </summary>
        private void Logon()
        {
            XDocument doc = XDocument.Load(String.Format("http://{0}/api.asp?cmd=logon&email={1}&password={2}", tbServer.Text, tbUser.Text, tbPassword.Text));
            FogBugzApiError error;
            if (doc.IsFogBugzError(out error)) error.Show(this);
            else if (doc.Descendants("token").Count<XElement>() == 1)
            {
                Settings.Default.Server = tbServer.Text;
                Settings.Default.User = tbUser.Text;
                Settings.Default.Password = tbPassword.Text;
                Settings.Default.Token = doc.Descendants("token").First<XElement>().Value;
                Settings.Default.Save();

                HideForm();
                tray.ShowBalloonTip(0, String.Format("Logged into {0}", Settings.Default.Server), "Now get crackin!", ToolTipIcon.Info);                
                updateTimer.Start();
                UpdateName();
            }
        }

        /// <summary>
        /// Updates the user's full name from FogBugz, which is used when listing cases.
        /// </summary>
        private void UpdateName()
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken("cmd=viewPerson"));
                FogBugzApiError error;
                if (doc.IsFogBugzError(out error)) error.Show(this);
                else
                {
                    IEnumerable<String> results = (
                        from p in doc.Descendants("person")
                        select p.Element("sFullName").Value);
                    if (results.Count() == 1)
                    {
                        Settings.Default.Name = results.ToArray()[0];
                        Settings.Default.Save();
                    }
                    else
                    {
                        MessageBox.Show("Invalid response from viewPerson. The application will now exit.", "Error updating from FogBugz", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
            }
        }
        
        /// <summary>
        /// Updates the internal project dictionary.
        /// </summary>
        private void UpdateProjects()
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken("cmd=listProjects"));
                FogBugzApiError error;
                if (doc.IsFogBugzError(out error)) error.Show(this);
                else
                {
                    projects = projects.FromKeyValuePairCollection<Int32, String>((
                        from p in doc.Descendants("project")
                        select new KeyValuePair<Int32, String>(Int32.Parse(p.Element("ixProject").Value), p.Element("sProject").Value)));
                }
            }
        }

        /// <summary>
        /// Updates the internal working case record.
        /// </summary>
        private void UpdateWorkingCase()
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken("cmd=listIntervals"));
                FogBugzApiError error;
                if (doc.IsFogBugzError(out error)) error.Show(this);
                else
                {
                    IEnumerable<Case> cases = (
                        from c in doc.Descendants("interval")
                        where c.Element("dtEnd").Value == ""
                        select new Case
                            {
                                Id = Int32.Parse(c.Element("ixBug").Value),
                                Title = c.Element("sTitle").Value
                            });

                    if (cases.Count() > 0) workingCase = cases.First();
                    else workingCase = null;

                    if (workingCase == null)
                    {
                        tray.Text = "Not working on anything.";
                        stopWorkToolStripMenuItem.Text = "&Stop Work";
                        stopWorkToolStripMenuItem.Enabled = false;
                    }
                    else
                    {
                        tray.Text = String.Format("Working on Case {0} - {1}.", workingCase.Id, workingCase.Title);
                        stopWorkToolStripMenuItem.Text = String.Format("&Stop Work on Case {0}", workingCase.Id);
                        stopWorkToolStripMenuItem.Enabled = true;
                    }
                }
            }
        }

        /// <summary>
        /// Updates the menu's lists of cases.
        /// </summary>
        private void UpdateCases()
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken(String.Format("cmd=search&q=assignedto:%22{0}%22%20status:active&cols=sTitle,ixProject,sFixFor", System.Web.HttpUtility.UrlEncode(Settings.Default.Name).Replace("+", "%20"))));
                FogBugzApiError error;
                if (doc.IsFogBugzError(out error)) error.Show(this);
                else
                {
                    IEnumerable<Case> cases = (
                        from c in doc.Descendants("case")
                        orderby Int32.Parse(c.Attribute("ixBug").Value) descending
                        select new Case
                        {
                            Id = Int32.Parse(c.Attribute("ixBug").Value),
                            Title = c.Element("sTitle").Value,
                            Project = projects[Int32.Parse(c.Element("ixProject").Value)],
                            FixFor = c.Element("sFixFor").Value
                        }
                        ).Take(10);

                    // Update the "Cases" menu.
                    casesToolStripMenuItem.DropDownItems.Clear();
                    foreach (Case cs in cases)
                    {
                        Boolean isSelected = workingCase == null ? false : workingCase.Id == cs.Id;
                        AddMenuItem(casesToolStripMenuItem, cs.Id, String.Format("{0} - {1}", cs.Id, cs.Title), Case_Click, isSelected);
                    }

                    // Update the "Projects" menu.
                    projectsToolStripMenuItem.DropDownItems.Clear();
                    // Create a menu item for each project.
                    foreach (IGrouping<String, Case> p in cases.GroupBy(c => c.Project))
                    {
                        ToolStripMenuItem projectMenu = new ToolStripMenuItem();
                        projectMenu.Text = p.Key;
                        projectsToolStripMenuItem.DropDownItems.Add(projectMenu);

                        // Create a menu item for each "Fix For" with the project.
                        foreach (IGrouping<String, Case> f in cases.Where(c => c.Project == p.Key).GroupBy(c => c.FixFor))
                        {
                            ToolStripMenuItem fixForMenu = new ToolStripMenuItem();
                            fixForMenu.Text = f.Key;
                            projectMenu.DropDownItems.Add(fixForMenu);

                            // Create a menu item for each case in the "Fix For".
                            foreach (Case cs in cases.Where(c => c.Project == p.Key && c.FixFor == f.Key))
                            {
                                Boolean isSelected = workingCase == null ? false : workingCase.Id == cs.Id;
                                AddMenuItem(fixForMenu, cs.Id, String.Format("{0} - {1}", cs.Id, cs.Title), Case_Click, isSelected);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Attempts to start the time tracking timer for the specified case.
        /// </summary>
        /// <param name="caseId">The FogBugz ID of the case to start work on.</param>
        /// <param name="error">The FogBugz error, if any occurs.</param>
        /// <returns>A Boolean value specifying whether the API call to start the time tracking timer succeeded.</returns>
        private Boolean TryStartWork(Int32 caseId, out FogBugzApiError error)
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken(String.Format("cmd=startWork&ixBug={0}", caseId)));
                if (doc.IsFogBugzError(out error))
                {
                    return false;
                }
                else
                {
                    UpdateWorkingCase();
                    UpdateCases();
                    tray.ShowBalloonTip(0, String.Format("Work started on Case {0}", caseId), workingCase.Title, ToolTipIcon.Info);
                    return true;
                }
            }
            else
            {
                error = null;
                return false;
            }
        }

        /// <summary>
        /// Starts the time tracking timer for the specified case.
        /// </summary>
        /// <param name="caseId">The FogBugz ID of the case to start work on.</param>
        private void StartWork(Int32 caseId)
        {
            if (IsLoggedIn)
            {
                FogBugzApiError error;
                TryStartWork(caseId, out error);
            }
        }

        /// <summary>
        /// Stops the time tracking timer for the case currently in progress.
        /// </summary>
        private void StopWork()
        {
            if (IsLoggedIn)
            {
                XDocument doc = XDocument.Load(GetCommandUrlWithToken("cmd=stopWork"));
                FogBugzApiError error;
                String title = String.Format("Work stopped on Case {0}", workingCase.Id);
                String text = workingCase.Title;
                if (doc.IsFogBugzError(out error)) error.Show(this);
                else
                {
                    tray.ShowBalloonTip(0, title, text, ToolTipIcon.Info);
                    UpdateWorkingCase();
                    UpdateCases();
                }
            }
        }

        #endregion

        #region [ Event Handlers ]

        /// <summary>
        /// Handles the click event of any case menu item.
        /// </summary>
        private void Case_Click(Object sender, EventArgs e)
        {
            ToolStripMenuItem menuItem = (ToolStripMenuItem)sender;

            // If the menu item was checked, stop the work.
            if (menuItem.Checked) StopWork();
            // Otherwise, start the work.
            else
            {
                Int32 caseId = (Int32)menuItem.Tag;
                FogBugzApiError error;
                if (!TryStartWork(caseId, out error)) error.Show(this);
            }
        }

        /// <summary>
        /// Handles the click event of the form's OK button.
        /// </summary>
        private void btnOk_Click(object sender, EventArgs e)
        {
            Logon();
            UpdateProjects();
            UpdateWorkingCase();
            UpdateCases();
        }

        /// <summary>
        /// Handles the click event of the "Log In" menu item.
        /// </summary>
        private void logInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm();
        }

        /// <summary>
        /// Handles the form's Closing event and forces the form to minimize
        /// instead of close when the "x" button is clicked on the toolbar.
        /// </summary>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                HideForm();
            }
        }

        /// <summary>
        /// Handles the click event of the Exit menu item and causes the
        /// application to exit.
        /// </summary>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Handles the click event of the Stop Work menu item.
        /// </summary>
        private void stopWorkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopWork();
        }

        /// <summary>
        /// Handles the Tick event of the form's updateTimer component, which
        /// periodically updates data from the FogBugz API.
        /// </summary>
        private void updateTimer_Tick(object sender, EventArgs e)
        {
            UpdateProjects();
            UpdateWorkingCase();
            UpdateCases();
        }

        /// <summary>
        /// Handles the KeyPress event of the manual case entry textbox in the menu.
        /// </summary>
        private void tbManualCase_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (Char)Keys.Enter)
            {
                Int32 caseId;
                if (Int32.TryParse(tbManualCase.Text, out caseId))
                {
                    menu.Visible = false;
                    tbManualCase.Text = String.Empty;
                    FogBugzApiError error;
                    if (!TryStartWork(caseId, out error))
                    {
                        tbManualCase.Text = caseId.ToString();
                        tbManualCase.SelectAll();
                    }
                }
                else MessageBox.Show(this, "Invalid Entry", "You value you entered is not a valid case number.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Handles the enter event of the manual case entry textbox, and causes
        /// the default text and styling in the textbox to vanish when the user enters it.
        /// </summary>
        private void tbManualCase_Enter(object sender, EventArgs e)
        {
            if (tbManualCase.Text == "Case #")
            {
                tbManualCase.Clear();
                tbManualCase.Font = new Font(tbManualCase.Font, FontStyle.Regular);
                tbManualCase.ForeColor = Color.FromKnownColor(KnownColor.WindowText);
            }
        }

        /// <summary>
        /// Handles the lost focus event of the manual case entry textbox and
        /// causes the default text and styling to reappear if the textbox is empty
        /// when the user leaves it.
        /// </summary>
        void TextBox_LostFocus(object sender, EventArgs e)
        {
            if (tbManualCase.Text == String.Empty)
            {
                tbManualCase.TextBox.Text = "Case #";
                tbManualCase.Font = new Font(tbManualCase.Font, FontStyle.Italic);
                tbManualCase.ForeColor = Color.Gray;
            }
        }

        /// <summary>
        /// Handles the click event of the menu. This is used as a workaround to
        /// cause the manual case entry textbox to lose focus when the user clicks
        /// elsewhere.
        /// </summary>
        private void menu_Click(object sender, EventArgs e)
        {
            menu.Focus();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateProjects();
            UpdateWorkingCase();
            UpdateCases();
        }

        #endregion

        #region [ Utility Methods ]

        /// <summary>
        /// Gets the FogBugz api request uri with the given command and the current session's authentication token.
        /// </summary>
        private String GetCommandUrlWithToken(String command)
        {
            return String.Format("http://{0}/api.asp?{1}&token={2}", Settings.Default.Server, command, Settings.Default.Token);
        }

        /// <summary>
        /// Adds a new menu item to the given menu item.
        /// </summary>
        /// <param name="parent">The menu item to add the child to.</param>
        /// <param name="tag">Any relevant data, such as an id.</param>
        /// <param name="text">The text of the menu item.</param>
        /// <param name="clickHandler">The click handler for the menu item.</param>
        /// <param name="isSelected">Whether or not the item should be checked.</param>
        private void AddMenuItem(ToolStripMenuItem parent, Object tag, String text, EventHandler clickHandler, Boolean isSelected)
        {
            ToolStripMenuItem menuItem = new ToolStripMenuItem();
            menuItem.Tag = tag;
            menuItem.Text = text;
            menuItem.Click += clickHandler;
            menuItem.Checked = isSelected;
            if (isSelected) menuItem.Font = new Font(menuItem.Font, FontStyle.Bold);
            parent.DropDownItems.Add(menuItem);
        }

        /// <summary>
        /// Normalizes and shows the logon form.
        /// </summary>
        private void ShowForm()
        {
            this.Visible = true;
            this.ShowInTaskbar = true;
            this.WindowState = FormWindowState.Normal;
        }

        /// <summary>
        /// Minimizes and hides the logon form.
        /// </summary>
        private void HideForm()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }

        /// <summary>
        /// Gets whether the user has a valid session.
        /// </summary>
        private Boolean IsLoggedIn
        {
            get { return !String.IsNullOrEmpty(Settings.Default.Token); }
        }

        #endregion

    }

    /// <summary>
    /// A container class for a FogBugz case.
    /// </summary>
    public class Case
    {
        public Int32 Id { get; set; }
        public String Title { get; set; }
        public String FixFor { get; set; }
        public String Project { get; set; }
    }

    /// <summary>
    /// A container class for FogBugz API errors.
    /// </summary>
    public class FogBugzApiError
    {
        public Int32 Code { get; set; }
        public String Message { get; set; }        
        public void Show(IWin32Window owner)
        {
            MessageBox.Show(owner, Message, String.Format("FogBugz API Error Code {0}", Code), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
