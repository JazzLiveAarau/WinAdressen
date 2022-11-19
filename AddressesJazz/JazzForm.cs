using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AdressesUtility;
using JazzVersion;

namespace AddressesJazz
{
    /// <summary>Main form for the JAZZ live AARAU address database
    /// <para>This is the Graphical User Interface to the application. Commands should not be executed in this class.</para>
    /// </summary>
    public partial class JazzForm : Form
    {
        #region Member variables

        /// <summary>Main class that executes most of the commands in this application</summary>
        private JazzMain m_main = null;

        /// <summary>Table that holds all addresses</summary>
        private ExcelUtil.Table m_table_addresses = null;

        /// <summary>Index for the current record</summary>
        private int m_row_index = 0;

        /// <summary>Search result record indices</summary>
        private int[] m_search_record_indices = null;

        /// <summary>Search result record display strings</summary>
        private string[] m_search_display_strings = null;

        /// <summary>Current version as string</summary>
        private string m_current_version = "Undefined";

        /// <summary>Flag telling if controls are being initialized</summary>
        bool m_is_initializing = false;

        #endregion // Member variables

        #region Create ToolTips
        private ToolTip m_tool_tip_application = new ToolTip();
        private ToolTip m_tool_tip_first_name = new ToolTip();
        private ToolTip m_tool_tip_family_name = new ToolTip();
        private ToolTip m_tool_tip_street = new ToolTip();
        private ToolTip m_tool_tip_street_number = new ToolTip();
        private ToolTip m_tool_tip_postal_code = new ToolTip();
        private ToolTip m_tool_tip_city = new ToolTip();
        private ToolTip m_tool_tip_email = new ToolTip();
        private ToolTip m_tool_tip_comment_one = new ToolTip();
        private ToolTip m_tool_tip_support = new ToolTip();
        private ToolTip m_tool_tip_delete = new ToolTip();
        private ToolTip m_tool_tip_add = new ToolTip();
        private ToolTip m_tool_tip_button_first_name = new ToolTip();
        private ToolTip m_tool_tip_button_family_name = new ToolTip();
        private ToolTip m_tool_tip_button_postal_code = new ToolTip();
        private ToolTip m_tool_tip_button_checkoutin = new ToolTip();
        private ToolTip m_tool_tip_previous_record = new ToolTip();
        private ToolTip m_tool_tip_next_record = new ToolTip();
        private ToolTip m_tool_tip_ouput_data = new ToolTip();
        private ToolTip m_tool_tip_season = new ToolTip();
        private ToolTip m_tool_tip_search = new ToolTip();
        private ToolTip m_tool_tip_search_hits = new ToolTip();
        private ToolTip m_tool_tip_status_message = new ToolTip();
        private ToolTip m_tool_tip_exit = new ToolTip();
        private ToolTip m_tool_tip_check_post = new ToolTip();
        private ToolTip m_tool_tip_check_newsletter = new ToolTip();
        private ToolTip m_tool_tip_check_sponsor = new ToolTip();
        private ToolTip m_tool_tip_search_results = new ToolTip();
        private ToolTip m_tool_tip_help = new ToolTip();
        private ToolTip m_tool_tip_update = new ToolTip();
        private ToolTip m_tool_tip_reset = new ToolTip();


        #endregion

        #region Constructor

        /// <summary>Constructor that initializes the form
        /// <para>1. Instance of JazzMain is created.</para>
        /// <para>2. Captions are set (call of _SetCaptions).</para>
        /// <para>3. Seasons controls are set (call of _SetSeasons).</para>
        /// <para>4. Combo Box for output files are set (call of _SetComboboxOutput).</para>
        /// <para>5. Gets the current address list from the server (calls JazzMainDownloadFile)</para>
        /// <para>6. Address Table is created. Call of JazzMain.CreateAddressesTable</para>
        /// <para>7. Text controls are set for the first record in the Table (call of SetControlsTexts).</para>
        /// <para>8. Search controls are initialized (call of _SetSearchResult).</para>
        /// <para>9. Status message (addresses downloaded) is set. Caption Checkout is set.</para>
        /// <para>10. Tool tips are set (call of _SetToolTips)</para>
        /// <para>11. Most controls are disabled since addresses not are checked out (call of _DisableControls).</para>
        /// </summary>
        public JazzForm()
        {
            InitializeComponent();

            // Necessary for CheckData.CheckNames
            this.m_textbox_first_name.Text = "Xxx";
            this.m_textbox_family_name.Text = "Yyy";

            if (AddressesJazzSettings.Default.BetaVersion)
            {
                this.m_label_version.Text = @"Beta Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    + "\nNur für Testzwecke";
            }
            else
            {
                this.m_label_version.Text = "";
            }

            this.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.JazzForm_MouseWheel);

            m_main = new JazzMain(this);

            string error_message = "";

            _SetCaptions();

            _SetSeasons();

            _SetComboboxOutput();

            if (!m_main.DownloadFile(out error_message))
            {
                Color color_error_bg = Color.Red;
                Color color_error_txt = Color.Blue;
                this.m_textbox_message.BackColor = color_error_bg;
                this.m_textbox_message.ForeColor = color_error_txt;
                this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgExitApplication;

                error_message = error_message + "\n" + AddressesJazzSettings.Default.MsgExitApplication;             
                MessageBox.Show(error_message);

                // Exception if this.Close() is called by the constructor 
                return;

                //if (!m_main.LocalFileExists(out error_message))
                //{
                    // Exit the application!
                    //MessageBox.Show(error_message);
                    //return;
                //}
            }

            if (!m_main.CreateAddressesTable(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            m_table_addresses = m_main.GetTable();

            m_row_index = 1; // First record
            SetControlsTexts(m_row_index);

            _SetSearchResult(); // Initialization

            this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgExcelFileDownload;

            this.m_button_checkoutin.Text = AddressesJazzSettings.Default.Caption_CheckOut;

            _VersionCheck();

            _SetToolTips();

            _DisableControls();

        } // Constructor

        #endregion // Constructor

        #region Reset table to backup

        /// <summary>Set new table (created by a backup file)
        /// <para>Set input Table instance equal to m_table_addresses</para>
        /// <para>TODO Delete no longer used instance</para>
        /// </summary>
        public void SetTable(ExcelUtil.Table i_table_addresses)
        {
            m_table_addresses = i_table_addresses;
        }

        #endregion // Reset table to backup

        #region Version check

        /// <summary>Checks if there is a new version available
        /// <para>Message will be written to the message control if available.</para>
        /// </summary>
        private void _VersionCheck()
        {
            VersionInput version_input = new VersionInput();

            version_input.FtpHost = AddressesJazzSettings.Default.FtpHost;
            version_input.FtpUser = AddressesJazzSettings.Default.FtpUser;
            version_input.FtpPassword = m_main.FtpPassword;
            version_input.ExeDirectory = System.Windows.Forms.Application.StartupPath;
            version_input.VersionString = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            version_input.ServerDirectory = @"/" + AddressesJazzSettings.Default.AdressenServerDir + @"/" +
                                                   AddressesJazzSettings.Default.LatestVersionInfoDir + @"/";
            version_input.LocalDirectory = AddressesJazzSettings.Default.LatestVersionInfoDir;
            
            string error_message = @"";
            if (!JazzVersion.VersionUtil.Init(version_input, out error_message))
            {
                m_label_version.Text = error_message;
                return;
            }

            string current_version_str = @"";
            if (!JazzVersion.VersionUtil.GetCurrentVersion(out current_version_str, out error_message))
            {
                this.m_textbox_message.Text = @"JazzForm._VersionCheck " + error_message;
                return;
            }

            m_current_version = " Version " + current_version_str;

            bool new_version = false;
            string new_version_str = @"";

            if (!JazzVersion.VersionUtil.NewVersionIsAvailable(out new_version, out new_version_str, out error_message))
            {
                this.m_textbox_message.Text = error_message;
                return;
            }

            this.m_label_version.Text = m_current_version;
            this.m_label_version.ForeColor = Color.White;

            if (new_version)
            {
                this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgNewVersionIsAvailable + new_version_str;
                this.m_textbox_message.BackColor = Color.Yellow;
            }

        } // _VersionCheck

        #endregion // Version check

        #region Tool tips

        /// <summary>Set tool tips
        /// <para>Note that there also are tool tips for Labels since disabled controls don't show tool tips</para>
        /// <para></para>
        /// </summary>
        private void _SetToolTips()
        {
            m_tool_tip_application.SetToolTip(this, AddressesJazzSettings.Default.ToolTipApplication + m_current_version);
            m_tool_tip_application.SetToolTip(this.m_picturebox_logo, AddressesJazzSettings.Default.ToolTipApplication + m_current_version);
            m_tool_tip_application.SetToolTip(this.m_panel_logo, AddressesJazzSettings.Default.ToolTipApplication + m_current_version);
            ToolTipUtil.SetDelays(ref m_tool_tip_application);

            m_tool_tip_first_name.SetToolTip(this.m_textbox_first_name, AddressesJazzSettings.Default.ToolTipTextBoxFirstName);
            m_tool_tip_first_name.SetToolTip(this.m_label_first_name, AddressesJazzSettings.Default.ToolTipTextBoxFirstName);
            ToolTipUtil.SetDelays(ref m_tool_tip_first_name);

            m_tool_tip_family_name.SetToolTip(this.m_textbox_family_name, AddressesJazzSettings.Default.ToolTipTextBoxFamyliName);
            m_tool_tip_family_name.SetToolTip(this.m_label_family_name, AddressesJazzSettings.Default.ToolTipTextBoxFamyliName);
            ToolTipUtil.SetDelays(ref m_tool_tip_family_name);

            m_tool_tip_street.SetToolTip(this.m_textbox_street, AddressesJazzSettings.Default.ToolTipTextBoxStreetName);
            m_tool_tip_street.SetToolTip(this.m_label_street, AddressesJazzSettings.Default.ToolTipTextBoxStreetName);
            ToolTipUtil.SetDelays(ref m_tool_tip_street);

            m_tool_tip_street_number.SetToolTip(this.m_textbox_street_number, AddressesJazzSettings.Default.ToolTipTextBoxHouseNumber);
            m_tool_tip_street_number.SetToolTip(this.m_label_street_number, AddressesJazzSettings.Default.ToolTipTextBoxHouseNumber);
            ToolTipUtil.SetDelays(ref m_tool_tip_street_number);

            m_tool_tip_postal_code.SetToolTip(this.m_textbox_postal_code, AddressesJazzSettings.Default.ToolTipTextBoxPostalCode);
            m_tool_tip_postal_code.SetToolTip(this.m_label_postal_code, AddressesJazzSettings.Default.ToolTipTextBoxPostalCode);
            ToolTipUtil.SetDelays(ref m_tool_tip_postal_code);

            m_tool_tip_city.SetToolTip(this.m_textbox_city, AddressesJazzSettings.Default.ToolTipTextBoxCityName);
            m_tool_tip_city.SetToolTip(this.m_label_city, AddressesJazzSettings.Default.ToolTipTextBoxCityName);
            ToolTipUtil.SetDelays(ref m_tool_tip_city);

            m_tool_tip_email.SetToolTip(this.m_textbox_email, AddressesJazzSettings.Default.ToolTipTextBoxEmailAddress);
            m_tool_tip_email.SetToolTip(this.m_label_email, AddressesJazzSettings.Default.ToolTipTextBoxEmailAddress);
            ToolTipUtil.SetDelays(ref m_tool_tip_email);

            m_tool_tip_comment_one.SetToolTip(m_textbox_comment_one, AddressesJazzSettings.Default.ToolTipTextBoxCommentOne);
            m_tool_tip_comment_one.SetToolTip(m_label_comment_one, AddressesJazzSettings.Default.ToolTipTextBoxCommentOne);
            ToolTipUtil.SetDelays(ref m_tool_tip_comment_one);

            m_tool_tip_support.SetToolTip(this.m_textbox_support, AddressesJazzSettings.Default.ToolTipTextBoxSupport);
            m_tool_tip_support.SetToolTip(this.m_label_support, AddressesJazzSettings.Default.ToolTipTextBoxSupport);
            ToolTipUtil.SetDelays(ref m_tool_tip_support);

            m_tool_tip_delete.SetToolTip(this.m_button_delete, AddressesJazzSettings.Default.ToolTipButtonDelete);
            ToolTipUtil.SetDelays(ref m_tool_tip_support);

            m_tool_tip_add.SetToolTip(this.m_button_add, AddressesJazzSettings.Default.ToolTipButtonAdd);
            ToolTipUtil.SetDelays(ref m_tool_tip_add);

            m_tool_tip_button_first_name.SetToolTip(this.m_button_first_name, AddressesJazzSettings.Default.ToolTipButtonSortFirstName);
            ToolTipUtil.SetDelays(ref m_tool_tip_button_first_name);

            m_tool_tip_button_family_name.SetToolTip(this.m_button_family_name, AddressesJazzSettings.Default.ToolTipButtonSortFamilyName);
            ToolTipUtil.SetDelays(ref m_tool_tip_button_family_name);

            m_tool_tip_button_postal_code.SetToolTip(this.m_button_postal_code, AddressesJazzSettings.Default.ToolTipButtonSortPostalCode);
            ToolTipUtil.SetDelays(ref m_tool_tip_button_postal_code);

            m_tool_tip_button_checkoutin.SetToolTip(this.m_button_checkoutin, AddressesJazzSettings.Default.ToolTipButtonCheckInOut);
            ToolTipUtil.SetDelays(ref m_tool_tip_button_checkoutin);

            m_tool_tip_check_post.SetToolTip(this.m_checkbox_post, AddressesJazzSettings.Default.ToolTipCheckPost);
            ToolTipUtil.SetDelays(ref m_tool_tip_check_post);

            m_tool_tip_check_newsletter.SetToolTip(this.m_checkbox_newsletter, AddressesJazzSettings.Default.ToolTipCheckNewsletter);
            ToolTipUtil.SetDelays(ref m_tool_tip_check_newsletter);

            m_tool_tip_check_sponsor.SetToolTip(this.m_checkbox_sponsor, AddressesJazzSettings.Default.ToolTipCheckSponsor);
            ToolTipUtil.SetDelays(ref m_tool_tip_check_sponsor);

            m_tool_tip_previous_record.SetToolTip(this.m_button_previous, AddressesJazzSettings.Default.ToolTipButtonPreviousRecord);
            ToolTipUtil.SetDelays(ref m_tool_tip_previous_record);

            m_tool_tip_next_record.SetToolTip(this.m_button_next, AddressesJazzSettings.Default.ToolTipButtonNextRecord);
            ToolTipUtil.SetDelays(ref m_tool_tip_next_record);

            m_tool_tip_ouput_data.SetToolTip(this.m_combobox_output, AddressesJazzSettings.Default.ToolTipOutputData);
            m_tool_tip_ouput_data.SetToolTip(this.m_label_output, AddressesJazzSettings.Default.ToolTipOutputData);
            ToolTipUtil.SetDelays(ref m_tool_tip_ouput_data);

            m_tool_tip_season.SetToolTip(this.m_combobox_season, AddressesJazzSettings.Default.ToolTipSeason);
            m_tool_tip_season.SetToolTip(this.m_label_season, AddressesJazzSettings.Default.ToolTipSeason);
            ToolTipUtil.SetDelays(ref m_tool_tip_season);

            m_tool_tip_search.SetToolTip(this.m_textbox_search, AddressesJazzSettings.Default.ToolTipSearchString);
            m_tool_tip_search.SetToolTip(this.m_label_search, AddressesJazzSettings.Default.ToolTipSearchString);
            ToolTipUtil.SetDelays(ref m_tool_tip_search);

            m_tool_tip_search_hits.SetToolTip(this.m_label_hits, AddressesJazzSettings.Default.ToolTipSearchStringHits);
            ToolTipUtil.SetDelays(ref m_tool_tip_search_hits);

            m_tool_tip_status_message.SetToolTip(this.m_textbox_message, AddressesJazzSettings.Default.ToolTipStatusMessage);
            ToolTipUtil.SetDelays(ref m_tool_tip_status_message);

            m_tool_tip_exit.SetToolTip(this.m_button_exit, AddressesJazzSettings.Default.ToolTipExitApplication);
            ToolTipUtil.SetDelays(ref m_tool_tip_exit);

            m_tool_tip_search_results.SetToolTip(this.m_combobox_search, AddressesJazzSettings.Default.ToolTipSearchResults);
            ToolTipUtil.SetDelays(ref m_tool_tip_search_results);

            m_tool_tip_help.SetToolTip(this.m_button_help, AddressesJazzSettings.Default.ToolTipButtonHelp + m_current_version);
            ToolTipUtil.SetDelays(ref m_tool_tip_help);
            m_tool_tip_update.SetToolTip(this.m_button_update, AddressesJazzSettings.Default.ToolTipButtonUpdate);
            ToolTipUtil.SetDelays(ref m_tool_tip_update);

            m_tool_tip_exit.SetToolTip(this.m_button_reset, AddressesJazzSettings.Default.ToolTipReset);
            ToolTipUtil.SetDelays(ref m_tool_tip_reset);

        } // _SetToolTips

        #endregion // Tool tips

        #region Set controls

        /// <summary>Set texts for all controls
        /// <para>Function GetFieldString in class Table is called for each field.</para>
        /// <para>Set default values for post (mail), newsletter and sponsor iv values not are set, i.e. for a new record</para>
        /// <para>Text field for the current season is set (call of _SetTextBoxSupport)</para>
        /// </summary>
        /// <param name="i_row_index">Index for the row (record) that shall be displayed</param>
        public void SetControlsTexts(int i_row_index)
        {
            if (m_table_addresses == null) return;

            string error_message = "";
            this.m_textbox_first_name.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_01, out error_message);
            this.m_textbox_family_name.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_02, out error_message);
            this.m_textbox_street.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_03, out error_message);
            this.m_textbox_street_number.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_04, out error_message);
            this.m_textbox_postal_code.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_05, out error_message);
            this.m_textbox_city.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_06, out error_message);
            this.m_textbox_email.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_07, out error_message);

            string check_box_post_value = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_08, out error_message);
            if (check_box_post_value.Length == 0)
            {
                check_box_post_value = "WAHR";
            }

            string check_box_newsletter_value = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_09, out error_message);
            if (check_box_newsletter_value.Length == 0)
            {
                check_box_newsletter_value = "WAHR";
            }

            string check_box_sponsor_value = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_10, out error_message);
            if (check_box_sponsor_value.Length == 0)
            {
                check_box_sponsor_value = "FALSCH";
            }

            _SetCheckBox(this.m_checkbox_post, check_box_post_value);
            _SetCheckBox(this.m_checkbox_newsletter, check_box_newsletter_value);
            _SetCheckBox(this.m_checkbox_sponsor, check_box_sponsor_value);

            this.m_textbox_comment_one.Text = m_table_addresses.GetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_11, out error_message);

            _SetTextBoxSupport(i_row_index);

            this.m_button_reset.Visible = false;

        } // SetControlsTexts

        /// <summary>Set text box support for the selected season</summary>
        private void _SetTextBoxSupport(int i_row_index)
        {
            string error_message = "";
            string column_name = "Beitrag-" + this.m_combobox_season.Text;
            string support_value_str = m_table_addresses.GetFieldString(i_row_index, column_name, out error_message);
            if (support_value_str == "0")
            {
                support_value_str = "";
            }
            this.m_textbox_support.Text = support_value_str;

            this.m_textbox_support.Enabled = true;
            string[] previous_versions = Season.GetPreviousSeasons();
            string selected_version = this.m_combobox_season.Text;
            for (int i_prev = 0; i_prev < previous_versions.Length; i_prev++)
            {
                string previous_version = previous_versions[i_prev];
                if (previous_version == selected_version)
                {
                    this.m_textbox_support.Enabled = false;
                }

            }

        } // _SetTextBoxSupport

        /// <summary>Set combobox seasons
        /// <para>Function GetAllSeasons in class Season is called.</para>
        /// </summary>
        private void _SetSeasons()
        {
            string[] displayed_seasons = Season.GetAllSeasons();

            for (int i_season = 0; i_season < displayed_seasons.Length; i_season++)
            {
                this.m_combobox_season.Items.Add(displayed_seasons[i_season]);
            }

            this.m_combobox_season.Text = Season.GetCurrentSeason();

        } // _SetSeasons

        /// <summary>Set combobox output</summary>
        private void _SetComboboxOutput()
        {
            m_is_initializing = true;

            string[] displayed_output = Output.GetAllOutput();

            for (int i_output = 0; i_output < displayed_output.Length; i_output++)
            {
                this.m_combobox_output.Items.Add(displayed_output[i_output]);
            }

            this.m_combobox_output.Text = displayed_output[0];

            m_is_initializing = false;

        } // _SetComboboxOutput

        /// <summary>Set captions (labels) for the controls
        /// <para>All captions are defined in the configuration file (class AddressesJazzSettings).</para>
        /// </summary>
        private void _SetCaptions()
        {

            this.m_label_first_name.Text = AddressesJazzSettings.Default.Caption_Record_01;
            this.m_label_family_name.Text = AddressesJazzSettings.Default.Caption_Record_02;
            this.m_label_street.Text = AddressesJazzSettings.Default.Caption_Record_03;
            this.m_label_street_number.Text = AddressesJazzSettings.Default.Caption_Record_04;
            this.m_label_postal_code.Text = AddressesJazzSettings.Default.Caption_Record_05;
            this.m_label_city.Text = AddressesJazzSettings.Default.Caption_Record_06;
            this.m_label_email.Text = AddressesJazzSettings.Default.Caption_Record_07;
            this.m_checkbox_post.Text = AddressesJazzSettings.Default.Caption_Record_08;
            this.m_checkbox_newsletter.Text = AddressesJazzSettings.Default.Caption_Record_09;
            this.m_checkbox_sponsor.Text = AddressesJazzSettings.Default.Caption_Record_10;
            this.m_label_comment_one.Text = AddressesJazzSettings.Default.Caption_Record_11;

            this.m_label_support.Text = AddressesJazzSettings.Default.Caption_Support;
            this.m_label_season.Text = AddressesJazzSettings.Default.Caption_Season;
            this.m_button_next.Text = AddressesJazzSettings.Default.Caption_Next;
            this.m_button_previous.Text = AddressesJazzSettings.Default.Caption_Previous;
            this.m_label_search.Text = AddressesJazzSettings.Default.Caption_Search;
            this.m_button_delete.Text = AddressesJazzSettings.Default.Caption_Delete;
            this.m_button_add.Text = AddressesJazzSettings.Default.Caption_Add;
            this.m_button_checkoutin.Text = AddressesJazzSettings.Default.Caption_CheckInOutUndefined;
            this.m_button_exit.Text = AddressesJazzSettings.Default.Caption_Exit;

        } // _SetCaptions

        /// <summary>Set check box to true or false</summary>
        private void _SetCheckBox(CheckBox i_checkbox, string i_str_boolean)
        {
            if (i_str_boolean == "WAHR")
            {
                i_checkbox.Checked = true;
            }
            else
            {
                i_checkbox.Checked = false;
            }

        } // _SetCheckBox

        /// <summary>Set check box to true or false</summary>
        private string _GetCheckBoxAsString(CheckBox i_checkbox)
        {
            string ret_str_boolean = "";

            if (i_checkbox.Checked)
            {
                ret_str_boolean = "WAHR";
            }
            else
            {
                ret_str_boolean = "FALSCH";
            }

            return ret_str_boolean;

        } // _GetCheckBoxAsString

        #endregion // Set controls

        #region Check record

        /// <summary>Check table record
        /// <para>Functions in class CheckData are called.</para>
        /// </summary>
        private bool _CheckTableRecord(out string o_error)
        {
            o_error = "";
            bool data_is_ok = true;

            string error_message = "";

            string first_name = this.m_textbox_first_name.Text.Trim();
            string family_name = this.m_textbox_family_name.Text.Trim();
            if (!CheckData.CheckNames(first_name, family_name, out error_message))
            {
                o_error = o_error + error_message + "\n";
                data_is_ok = false;
            }


            string mail_address = this.m_textbox_email.Text.Trim();
            if (!CheckData.CheckEmailAddress(mail_address, out error_message))
            {
                o_error = o_error + error_message + "\n";
                data_is_ok = false;
            }

            if (this.m_checkbox_post.Checked)
            {
                string street = this.m_textbox_street.Text.Trim();
                string street_number = this.m_textbox_street_number.Text.Trim();
                string postal_code = this.m_textbox_postal_code.Text.Trim();
                string city = this.m_textbox_city.Text.Trim();
                if (!CheckData.CheckMailAddress(street, street_number, postal_code, city, out error_message))
                {
                    o_error = o_error + error_message + "\n";
                    data_is_ok = false;
                }
            }

            if (this.m_checkbox_newsletter.Checked)
            {
                if (!CheckData.EmailAddressExists(mail_address, out error_message))
                {
                    o_error = o_error + error_message + "\n";
                    data_is_ok = false;
                }
            }

            if (data_is_ok == true)
            {
                return true;
            }
            else
            {
                return false;
            }

        } // _CheckTableRecord

        #endregion // Check record

        #region Set table record

        /// <summary>Set table record, i.e. write the data to Table m_table_addresses
        /// <para>(Changes that the user may have done are saved).</para>
        /// <para>Function SetFieldString in class Table is called for each field.</para>
        /// <para>Note that when there is no support the displayed field is empty, but in the Table it must be set to 0.</para>
        /// </summary>
        private void _SetTableRecord(int i_row_index)
        {
            if (m_table_addresses == null) return;

            string error_message = "";
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_01, this.m_textbox_first_name.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_02, this.m_textbox_family_name.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_03, this.m_textbox_street.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_04, this.m_textbox_street_number.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_05, this.m_textbox_postal_code.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_06, this.m_textbox_city.Text.Trim(), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_07, this.m_textbox_email.Text.Trim(), out error_message)) return;

            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_08, _GetCheckBoxAsString(this.m_checkbox_post), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_09, _GetCheckBoxAsString(this.m_checkbox_newsletter), out error_message)) return;
            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_10, _GetCheckBoxAsString(this.m_checkbox_sponsor), out error_message)) return;

            if (!m_table_addresses.SetFieldString(i_row_index, AddressesJazzSettings.Default.Name_Record_11, this.m_textbox_comment_one.Text.Trim(), out error_message)) return;
            
            // Column name for the selected season
            string column_name = "Beitrag-" + this.m_combobox_season.Text;
            // No support must be saved as string 0 in the Table
            string support_value = this.m_textbox_support.Text.Trim();
            if (support_value == "")
            {
                support_value = "0";
            }

            if (!m_table_addresses.SetFieldString(i_row_index, column_name, support_value, out error_message)) return;

        }

        #endregion // Set table record

        #region Disable and enable controls

        /// <summary>Disable controls (when addresses not are checked out)
        /// <para>Colours for the disabled controls are set to LightGoldenrodYellow</para>
        /// </summary>
        private void _DisableControls()
        {
            bool ctrl_disable = false;
            Color color_disable = Color.LightGoldenrodYellow;

            this.m_button_checkoutin.BackColor = System.Drawing.SystemColors.Control;

            this.m_textbox_first_name.Enabled = ctrl_disable;
            this.m_textbox_first_name.BackColor = color_disable;
            this.m_textbox_family_name.Enabled = ctrl_disable;
            this.m_textbox_family_name.BackColor = color_disable;
            this.m_textbox_street.Enabled = ctrl_disable;
            this.m_textbox_street.BackColor = color_disable;
            this.m_textbox_street_number.Enabled = ctrl_disable;
            this.m_textbox_street_number.BackColor = color_disable;
            this.m_textbox_postal_code.Enabled = ctrl_disable;
            this.m_textbox_postal_code.BackColor = color_disable;
            this.m_textbox_city.Enabled = ctrl_disable;
            this.m_textbox_city.BackColor = color_disable;
            this.m_textbox_email.Enabled = ctrl_disable;
            this.m_textbox_email.BackColor = color_disable;
            this.m_checkbox_post.Enabled = ctrl_disable;
            this.m_checkbox_post.BackColor = color_disable;
            this.m_checkbox_newsletter.Enabled = ctrl_disable;
            this.m_checkbox_newsletter.BackColor = color_disable;
            this.m_checkbox_sponsor.Enabled = ctrl_disable;
            this.m_checkbox_sponsor.BackColor = color_disable;
            this.m_textbox_comment_one.Enabled = ctrl_disable;
            this.m_textbox_comment_one.BackColor = color_disable;
            this.m_textbox_support.Enabled = ctrl_disable;
            this.m_textbox_support.BackColor = color_disable;
            this.m_button_delete.Enabled = ctrl_disable;
            this.m_button_delete.BackColor = color_disable;
            this.m_button_add.Enabled = ctrl_disable;
            this.m_button_add.BackColor = color_disable;
            this.m_button_first_name.Enabled = ctrl_disable;
            this.m_button_first_name.BackColor = color_disable;
            this.m_button_family_name.Enabled = ctrl_disable;
            this.m_button_family_name.BackColor = color_disable;
            this.m_button_postal_code.Enabled = ctrl_disable;
            this.m_button_postal_code.BackColor = color_disable;
            this.m_button_reset.Enabled = ctrl_disable;

            this.m_combobox_search.BackColor = color_disable;
        }

        /// <summary>Enable controls (when addresses are checked out)
        /// <para>Color for the (Checkout)/Save button is set to OrangeRed.</para>
        /// </summary>
        private void _EnableControls()
        {
            bool ctrl_enable = true;
            Color color_enable = System.Drawing.SystemColors.Window;

            this.m_button_checkoutin.BackColor = Color.OrangeRed;

            this.m_textbox_first_name.Enabled = ctrl_enable;
            this.m_textbox_first_name.BackColor = color_enable;
            this.m_textbox_family_name.Enabled = ctrl_enable;
            this.m_textbox_family_name.BackColor = color_enable;
            this.m_textbox_street.Enabled = ctrl_enable;
            this.m_textbox_street.BackColor = color_enable;
            this.m_textbox_street_number.Enabled = ctrl_enable;
            this.m_textbox_street_number.BackColor = color_enable;
            this.m_textbox_postal_code.Enabled = ctrl_enable;
            this.m_textbox_postal_code.BackColor = color_enable;
            this.m_textbox_city.Enabled = ctrl_enable;
            this.m_textbox_city.BackColor = color_enable;
            this.m_textbox_email.Enabled = ctrl_enable;
            this.m_textbox_email.BackColor = color_enable;
            this.m_checkbox_post.Enabled = ctrl_enable;
            this.m_checkbox_post.BackColor = color_enable;
            this.m_checkbox_newsletter.Enabled = ctrl_enable;
            this.m_checkbox_newsletter.BackColor = color_enable;
            this.m_checkbox_sponsor.Enabled = ctrl_enable;
            this.m_checkbox_sponsor.BackColor = color_enable;
            this.m_textbox_comment_one.Enabled = ctrl_enable;
            this.m_textbox_comment_one.BackColor = color_enable;
            this.m_textbox_support.Enabled = ctrl_enable;
            this.m_textbox_support.BackColor = color_enable;
            this.m_button_delete.Enabled = ctrl_enable;
            this.m_button_delete.BackColor = color_enable;
            this.m_button_add.Enabled = ctrl_enable;
            this.m_button_add.BackColor = color_enable;
            this.m_button_first_name.Enabled = ctrl_enable;
            this.m_button_first_name.BackColor = color_enable;
            this.m_button_family_name.Enabled = ctrl_enable;
            this.m_button_family_name.BackColor = color_enable;
            this.m_button_postal_code.Enabled = ctrl_enable;
            this.m_button_postal_code.BackColor = color_enable;
            this.m_button_reset.Enabled = ctrl_enable;

            Color color_disable = Color.LightGoldenrodYellow;
            this.m_combobox_search.BackColor = color_disable;
        }

        #endregion // Disable and enable controls

        #region Sorting

        /// <summary>Sort with postal code
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Sort with function JazzMain.Sort</para>
        /// <para>Display first record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_postal_code_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            string record_name = AddressesJazzSettings.Default.Name_Record_05;
            if (!m_main.Sort(record_name, out error_message))
            {
                MessageBox.Show(error_message);
            }

            m_row_index = 1;
            SetControlsTexts(m_row_index);

        } // m_button_postal_code_Click

        /// <summary>Sort with first name
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Sort with function JazzMain.Sort</para>
        /// <para>Display first record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_first_name_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            string record_name = AddressesJazzSettings.Default.Name_Record_01;
            if (!m_main.Sort(record_name, out error_message))
            {
                MessageBox.Show(error_message);
            }

            m_row_index = 1;
            SetControlsTexts(m_row_index);

        } // m_button_first_name_Click

        /// <summary>Sort with family name
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Sort with function JazzMain.Sort</para>
        /// <para>Display first record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_family_name_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            string record_name = AddressesJazzSettings.Default.Name_Record_02;
            if (!m_main.Sort(record_name, out error_message))
            {
                MessageBox.Show(error_message);
            }

            m_row_index = 1;
            SetControlsTexts(m_row_index);

        } // m_button_family_name_Click

        #endregion // Sorting

        #region Next and previous record

        /// <summary>Mouse wheel was rotated
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Convert rotation to new current record index in Table</para>
        /// <para>Display current record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void JazzForm_MouseWheel(object sender, MouseEventArgs e)
        {
            if (m_table_addresses == null) return;

            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            int wheel_notches = e.Delta / 120;

            if (wheel_notches > 0)
            {
                if (m_row_index >= 2)
                {
                    m_row_index = m_row_index - 1;
                    SetControlsTexts(m_row_index);
                }
            }
            else
            {
                if (m_row_index <= m_table_addresses.NumberRows - 2)
                {
                    m_row_index = m_row_index + 1;
                    SetControlsTexts(m_row_index);
                }
            }

        } // JazzForm_MouseWheel

        /// <summary>Display previous row (record)
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Change current index to previos index</para>
        /// <para>Display current record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_previous_Click(object sender, EventArgs e)
        {
            if (m_table_addresses == null) return;

            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            if (m_row_index >= 2)
            {
                m_row_index = m_row_index - 1;
                SetControlsTexts(m_row_index);
            }

        } // m_button_previous_Click

        /// <summary>Display next row (record)
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Change current index to next index</para>
        /// <para>Display current record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_next_Click(object sender, EventArgs e)
        {
            if (m_table_addresses == null) return;

            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            if (m_row_index <= m_table_addresses.NumberRows - 2)
            {
                m_row_index = m_row_index + 1;
                SetControlsTexts(m_row_index);
            }

        } // m_button_next_Click

        #endregion // Next and previous record

        #region Delete record

        /// <summary>Delete a record
        /// <para>Call of function JazzMain.RemoveRow</para>
        /// <para>Display first record in the Table (call of SetControlsTexts).</para>
        /// </summary>
        private void m_button_delete_Click(object sender, EventArgs e)
        {
            _SetTableRecord(m_row_index);

            string error_message = "";
            if (m_main.RemoveRow(m_row_index, out error_message))
            {
                m_row_index = 1;
                SetControlsTexts(m_row_index);
            }

        } // m_button_delete_Click

        #endregion // Delete record

        #region Add record

        /// <summary>Add an address record
        /// <para>Current (displayed) record must be checked before setting (saving) it. (call of _CheckTableRecord)</para>
        /// <para>Set record (call of _SetTableRecord)</para>
        /// <para>Call of JazzMain.AppendEmptyRow that adds a record. Record index is returned</para>
        /// <para>Set support to 0 for all seasons (call of JazzMain.SetSupportToZeroForRow)</para>
        /// <para>Display new record in the Table (call of SetControlsTexts)</para>
        /// </summary>
        private void m_button_add_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            int empty_row_index = -1;
            if (m_main.AppendEmptyRow(out empty_row_index, out error_message))
            {
                if (m_main.SetSupportToZeroForRow(empty_row_index, out error_message))
                {
                    m_row_index = empty_row_index;
                    SetControlsTexts(m_row_index);
                }
            }

        } // m_button_add_Click

        #endregion // Add record

        #region Exit application

        /// <summary>Exit application. 
        /// <para>1. Handle the case when addresses are checked out. Call of _ExitWhenAddressesAreCheckedOut.</para>
        /// <para>2. Remove all temporary used files. Call of JazzMain.RemoveAllTemporaryUsedFiles.</para>
        /// </summary>
        private void m_button_exit_Click(object sender, EventArgs e)
        {
            if (this.m_main.AddressesCheckedOut)
            {
                if (!_ExitWhenAddressesAreCheckedOut())
                    return;
            }

            string error_message = "";
            if (!this.m_main.RemoveAllTemporaryUsedFiles(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            this.Close();

        } // m_button_exit_Click

        /// <summary>Main form is closing when the user has killed the main form with the Cancel button.
        /// <para>Handles the case when addresses are checked out. Addresses (changes) will not be saved.</para>
        /// <para>1. If addresses are are checked out call of JazzMain.UploadCheckOutInFileButNotAddresses.</para>
        /// <para>2. Delete all temporary used files. Call of JazzMain.RemoveAllTemporaryUsedFiles</para>
        /// </summary>
        private void JazzForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            string error_message = "";

            // Handles the case when user killed the main form with Cancel
            if (this.m_main.AddressesCheckedOut)
            {
                error_message = AddressesJazzSettings.Default.ErrMsgCancelWithoutSave;
                MessageBox.Show(error_message);

                if (!m_main.UploadCheckOutInFileButNotAddresses(out error_message))
                {
                    MessageBox.Show(error_message);
                    return;
                }
            }

            if (!this.m_main.RemoveAllTemporaryUsedFiles(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

        } // JazzForm_FormClosing

        /// <summary>Handles the case when the user exits the application and the addresses are checked out</summary>
        private bool _ExitWhenAddressesAreCheckedOut()
        {
            if (m_main.AddressesCheckedOut == false)
                return true;

            string error_message = "";
            string save_message = "";
            string caption_message = "";
            if (!m_main.GetSaveAddressesMessage(out save_message, out caption_message, out error_message))
            {
                MessageBox.Show(error_message);
                return false;
            }

            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            result = MessageBox.Show(save_message, caption_message, buttons);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                if (!_CheckTableRecord(out error_message))
                {
                    MessageBox.Show(error_message);
                    return false;
                }
                _SetTableRecord(m_row_index);

                if (!_CheckInAddresses())
                {
                    return false;
                }
            }
            else
            {
                // Checkin without upload of addresses must be made
                if (!m_main.UploadCheckOutInFileButNotAddresses(out error_message))
                {
                    MessageBox.Show(error_message);
                    return false;
                }
            }

            return true;

        } // _ExitWhenAddressesAreCheckedOut

        #endregion // Exit application

        #region Checkin and checkout

        /// <summary>Check in the CSV file with addresses
        /// <para>Call of JazzMain.CreateAndUploadFiles</para>
        /// <para>Change button Save to caption Checkout</para>
        /// <para>Set status message</para>
        /// <para>Disable controls (call of _DisableControls).</para>
        /// </summary>
        private bool _CheckInAddresses()
        {
            // Select functions depending on Checkin/Checkout status
            // Check if changes have been made ....

            string error_message = "";


            if (!m_main.CreateAndUploadFiles(out error_message))
            {
                MessageBox.Show(error_message);
                return false;
            }

            this.m_button_checkoutin.Text = AddressesJazzSettings.Default.Caption_CheckOut;

            this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgAddressesAreCheckedIn;

            _DisableControls();

            return true;

        } // _CheckInAddresses

        /// <summary>Check out file with addresses
        /// <para>1. Update the table with current address data from the server. Call of _UpdateTableWithAddressesFromServer.</para>
        /// <para>2. Checkout. Call of JazzMain.CheckOutAddresses</para>
        /// <para>   Show error message if addresses are checked out by somebody else</para>
        /// <para>   For this case give the user the possibility to anyhow checkout addresses</para>
        /// <para>   If the user wants this call JazzMain.ForceCheckOutAddresses</para>
        /// <para>3. Change button caption from Checkout to Save</para>
        /// <para>4. Display status message</para>
        /// <para>5. Enable the controls. Call of _EnableControls.</para>
        /// </summary>
        private bool _CheckOutAddresses()
        {
            string error_message = "";

            if (!_UpdateTableWithAddressesFromServer(out error_message))
            {
                MessageBox.Show(error_message);

                return false;
            }

            if (!m_main.CheckOutAddresses(out error_message))
            {
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;
                result = MessageBox.Show(error_message, "", buttons);

                if (result == System.Windows.Forms.DialogResult.No)
                {
                    return false;
                }
                else
                {
                    if (!m_main.ForceCheckOutAddresses(out error_message))
                    {
                        MessageBox.Show(error_message);
                        return false;
                    }
                }
            }

            this.m_button_checkoutin.Text = AddressesJazzSettings.Default.Caption_CheckIn;

            this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgAddressesAreCheckedOut;

            _EnableControls();

            return true;

        } // _CheckOutAddresses

        /// <summary>Update table with addresses data from the server
        /// <para>This function should be called just before a checkout is made. Other persons may have changed data.</para>
        /// <para>1. Function shall only be called when addresses not are checked out. Call of JazzMain.AddressesCheckedOut.</para>
        /// <para>2. Get names for current record. Calls of Table.GetFieldString</para>
        /// <para>3. Download the csv address file from the server. Call of JazzMain.DownloadFile.</para>
        /// <para>4. Create the table with the downloaded file. Call of JazzMain.CreateAddressesTable.</para>
        /// <para>5. Set the current record to one (1) if names not are equal, i.e. somebody else has added rows.</para>
        /// <para>6. Set the controls for the current record. Call of SetControlsTexts.</para>
        /// <para></para>
        /// </summary>
        private bool _UpdateTableWithAddressesFromServer(out string o_error)
        {
            this.m_textbox_message.Text = "";

            o_error = "";
            if (m_main.AddressesCheckedOut)
            {
                o_error = "JazzForm._UpdateTableWithAddressesFromServer Addresses is already checked out.";
                return false;
            }

            string error_message = "";
            string first_name = m_table_addresses.GetFieldString(m_row_index, AddressesJazzSettings.Default.Name_Record_01, out error_message);
            string family_name = m_table_addresses.GetFieldString(m_row_index, AddressesJazzSettings.Default.Name_Record_02, out error_message);

            if (!m_main.DownloadFile(out o_error))
            {
                o_error = "JazzForm._UpdateTableWithAddressesFromServer JazzMain.DownloadFile failed " + o_error;
                this.m_textbox_message.Text = o_error;
                return false;
            }

            if (!m_main.CreateAddressesTable(out o_error))
            {
                o_error = "JazzForm._UpdateTableWithAddressesFromServer JazzMain.DownloadFile failed " + o_error;
                this.m_textbox_message.Text = o_error;
                return false;
            }

            m_table_addresses = m_main.GetTable();

            string first_name_new_table = m_table_addresses.GetFieldString(m_row_index, AddressesJazzSettings.Default.Name_Record_01, out error_message);
            string family_name_new_table = m_table_addresses.GetFieldString(m_row_index, AddressesJazzSettings.Default.Name_Record_02, out error_message);

            if (!first_name_new_table.Equals(first_name) || !family_name_new_table.Equals(family_name))
            {
                m_row_index = 1;
            }

            SetControlsTexts(m_row_index);

            this.m_textbox_message.Text = AddressesJazzSettings.Default.MsgExcelFileDownload;

            return true;

        } // _UpdateTableWithAddressesFromServer


        /// <summary>Checkin or checkout file with addresses</summary>
        private void m_button_checkinout_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            _SetTableRecord(m_row_index);

            if (this.m_main.AddressesCheckedOut)
            {
                bool b_checkin = _CheckInAddresses();
            }
            else
            {
                bool b_checkout = _CheckOutAddresses();
            }

        } // m_button_checkinout_Click

        #endregion // Checkin and checkout

        #region Search

        /// <summary>Search string was changed
        /// <para>Call function JazzMain.Search</para>
        /// <para>Set seach result (call of _SetSearchResult).</para>
        /// <para>Set table record (call of _SetTableRecord).</para>
        /// <para>Set the selected record (call of _SetSelectedSearchRecord).</para>
        /// <para></para>
        /// </summary>
        private void m_textbox_search_TextChanged(object sender, EventArgs e)
        {
            string error_message = "";
            if (m_main.Search(this.m_textbox_search.Text, out m_search_display_strings, out m_search_record_indices, out error_message))
            {
                _SetSearchResult();

                _SetTableRecord(m_row_index);

                int search_index_record = 0;
                _SetSelectedSearchRecord(search_index_record);
            }
            else
            {
                MessageBox.Show(error_message);
            }

        } // m_textbox_search_TextChanged

        /// <summary>Set search result</summary>
        private void _SetSearchResult()
        {
            this.m_combobox_search.Items.Clear();
            this.m_combobox_search.Text = "";
            this.m_label_hits.Text = "0";

            if (m_search_display_strings == null || m_search_record_indices == null ||
                m_search_display_strings.Length == 0 || m_search_record_indices.Length == 0)
            {
                return;
            }

            for (int i_found = 0; i_found < m_search_display_strings.Length; i_found++)
            {
                string display_string = m_search_display_strings[i_found];

                this.m_combobox_search.Items.Add(display_string);

                if (0 == i_found)
                {
                    this.m_combobox_search.Text = display_string;
                }

                this.m_label_hits.Text = m_search_display_strings.Length.ToString();
            }

        } // _SetSearchResult

        /// <summary>User has selected another search record</summary>
        private void m_combobox_search_SelectedIndexChanged(object sender, EventArgs e)
        {
            string error_message = "";
            if (_CheckTableRecord(out error_message))
            {
                _SetTableRecord(m_row_index);
            }
            else
            {
                error_message = error_message + "\n" + AddressesJazzSettings.Default.ErrMsgRecordNotSaved;
                MessageBox.Show(error_message);
            }

            int search_index_record = this.m_combobox_search.SelectedIndex;

            _SetSelectedSearchRecord(search_index_record);

        } // m_combobox_search_SelectedIndexChanged

        /// <summary>Set the selected search record</summary>
        private void _SetSelectedSearchRecord(int i_search_index_record)
        {
            if (m_search_display_strings == null || m_search_record_indices == null ||
                m_search_display_strings.Length == 0 || m_search_record_indices.Length == 0)
            {
                return;
            }

            if (i_search_index_record < 0 || i_search_index_record >= m_search_record_indices.Length)
            {
                return;
            }

            m_row_index = m_search_record_indices[i_search_index_record];
            SetControlsTexts(m_row_index);

        } // _SetSelectedSearchRecord

        #endregion // Search

        #region User changed address record

        /// <summary>Remove invalid characters from input</summary>
        private void _RemoveInvalidCharsFromInput(TextBox i_text_box)
        {
            string input_string = i_text_box.Text;
            bool b_mod = false;
            string mod_string = StringUtil.RemoveInvalidCharsForCsv(input_string, out b_mod);

            if (b_mod)
            {
                string error_message = AddressesJazzSettings.Default.ErrMsgNotValidCharsHaveBeenRemoved + "\n"
                   + input_string + "\n" + mod_string;
                i_text_box.Text = mod_string;
                MessageBox.Show(error_message);
            }
        }

        /// <summary>Remove all characters except numbers</summary>
        private void _RemoveAllCharsButNumbers(TextBox i_text_box)
        {
            string input_string = i_text_box.Text;
            bool b_mod = false;
            string mod_string = StringUtil.RemoveAllCharsButNumbers(input_string, out b_mod);

            if (b_mod)
            {
                string error_message = AddressesJazzSettings.Default.ErrMsgAllCharsExceptNumbersHaveBeenRemoved + "\n" 
                    + input_string + "\n" + mod_string;
                MessageBox.Show(error_message);

                i_text_box.Text = mod_string;
            }
        }

        /// <summary>User changed the first name.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_first_name_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_first_name);
        }

        /// <summary>User changed the family name.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_family_name_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_family_name);
        }

        /// <summary>User changed the Email addresse.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_email_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_email);
        }

        /// <summary>User changed the post flag.
        /// <para>Do nothing</para>
        /// </summary>
        private void m_checkbox_post_CheckedChanged(object sender, EventArgs e)
        {
            // Do nothing
        }

        /// <summary>User changed the newsletter flag.</summary>
        private void m_checkbox_newsletter_CheckedChanged(object sender, EventArgs e)
        {
            // Do nothing
        }

        /// <summary>User changed the sponsor flag.
        /// <para>Do nothing</para>
        /// </summary>
        private void m_checkbox_sponsor_CheckedChanged(object sender, EventArgs e)
        {
            // Do nothing
        }

        /// <summary>User changed the street name.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_street_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_street);
        }

        /// <summary>User changed the street number.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_street_number_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_street_number);
        }

        /// <summary>User changed the postal code.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_postal_code_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_postal_code);
        }

        /// <summary>User changed the city name.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_city_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_city);
        }

        /// <summary>User changed the comment.
        /// <para>Call of _RemoveInvalidCharsFromInput</para>
        /// </summary>
        private void m_textbox_comment_one_TextChanged(object sender, EventArgs e)
        {
            _RemoveInvalidCharsFromInput(this.m_textbox_comment_one);
        }

        /// <summary>Keydown comment.
        /// <para>Do nothing</para>
        /// </summary>
        private void m_textbox_comment_one_KeyDown(object sender, KeyEventArgs e)
        {
            // Not used, only tested
        }
        /// <summary>Keypress comment.
        /// <para>Do nothing</para>
        /// </summary>
        private void m_textbox_comment_one_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Not used, only tested
        }

        /// <summary>User changed season.
        /// <para>Call of SetControlsTexts</para>
        /// </summary>
        private void m_combobox_season_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetControlsTexts(m_row_index);
        }

        /// <summary>User changed the support sum.
        /// <para>Call of _RemoveAllCharsButNumbers</para>
        /// </summary>
        private void m_textbox_support_TextChanged(object sender, EventArgs e)
        {
            _RemoveAllCharsButNumbers(this.m_textbox_support);
        }

        #endregion // User changed address record

        #region Output file

        /// <summary>User has set an output request.
        /// <para>Requests are executed by function ExecuteRequest in class Output</para>
        /// <para></para>
        /// </summary>
        private void m_combobox_output_SelectedIndexChanged(object sender, EventArgs e)
        {
            string error_message = "";
            if (!_CheckTableRecord(out error_message))
            {
                this.m_combobox_output.Text = Output.GetHeaderItem(); 
                MessageBox.Show(error_message);
                return;
            }

            string selected_inner_text = this.m_combobox_output.Text;
            int selected_index = this.m_combobox_output.SelectedIndex;
            string season_column_name = "Beitrag-" + this.m_combobox_season.Text;

            string output_file_name = "";
            if (selected_index != 0)
            {
                output_file_name = m_main.GetOutputFileName(selected_inner_text);
            }

            if (!Output.ExecuteRequest(this.m_main.GetTable(), season_column_name, selected_index, selected_inner_text, output_file_name, out error_message))
            {
                MessageBox.Show(error_message);
            }

            this.m_combobox_output.Text = Output.GetHeaderItem();

            if (!m_is_initializing)
            {
                AlwaysRecreateTableAfterOutput(out error_message);
            }
            
        } // m_combobox_output_SelectedIndexChanged

        /// <summary>TODO This should not be necessary
        /// <para>Obviously are the output functions changing the table</para>
        /// </summary>
        private bool AlwaysRecreateTableAfterOutput(out string o_error)
        {
            o_error = "";

            // QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ
            // TODO The table has to be recreated, while it has changed .... Why has it been changed ????????????????????????????
            if (!m_main.CreateAddressesTable(out o_error))
            {
                return false;
            }

            m_table_addresses = m_main.GetTable();

            return true;
        }

        #endregion // Output file

        #region Help and download new version

        /// <summary>User clicked the help button
        /// <para>Show dialog FormHelp</para>
        /// </summary>
        private void m_button_help_Click(object sender, EventArgs e)
        {
            FormHelp help_form = new FormHelp(JazzMain.m_exe_directory);
            help_form.Owner = this;
            help_form.ShowDialog();
        }

        /// <summary>User clicked the download button
        /// <para>Call of JazzMain.DownloadNewVersion</para>
        /// </summary>
        private void m_button_update_Click(object sender, EventArgs e)
        {
            string error_message = "";
            if (!m_main.DownloadNewVersion(out error_message))
            {
                MessageBox.Show(error_message);
            }
            else
            {
                MessageBox.Show(AddressesJazzSettings.Default.MsgNewVersionDownload);
            }
        }

        #endregion // Help and download new version

        #region Reset

        /// <summary>User clicked the reset button
        /// <para>Show dialog ResetForm</para>
        /// </summary>
        private void m_button_reset_Click(object sender, EventArgs e)
        {
            if (!this.m_main.AddressesCheckedOut)
            {
                MessageBox.Show(AddressesJazzSettings.Default.ErrMsgAddressesMustBeCheckedOutForReset);
                return;
            }

            ResetForm reset_form = new ResetForm(this, m_main);
            reset_form.Owner = this;
            reset_form.ShowDialog();

        } // m_button_reset_Click

        #endregion // Reset

    } // JazzForm

} // namespace
