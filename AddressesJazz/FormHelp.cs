using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AdressesUtility;

namespace AddressesJazz
{
    /// <summary>Shows help</summary>
    public partial class FormHelp : Form
    {
        private ToolTip m_tool_tip_help = new ToolTip();
        private ToolTip m_tool_tip_help_text = new ToolTip();
        private ToolTip m_tool_tip_help_exit = new ToolTip();

        private string m_test_protokoll_name = "";

        /// <summary>Constructor that displays the help file</summary>
        public FormHelp(string i_exe_dir)
        {
            InitializeComponent();

            this.Text = AddressesJazzSettings.Default.GuiHelpDialogTitle;
            this.m_button_help_exit.Text = AddressesJazzSettings.Default.GuiHelpDialogExit;

            m_tool_tip_help.SetToolTip(this, AddressesJazzSettings.Default.ToolTipButtonHelp);
            ToolTipUtil.SetDelays(ref m_tool_tip_help);
            m_tool_tip_help_text.SetToolTip(this.m_rich_text_box_help, AddressesJazzSettings.Default.ToolTipHelpTextBox);
            ToolTipUtil.SetDelays(ref m_tool_tip_help_text);
            m_tool_tip_help_exit.SetToolTip(this.m_button_help_exit, AddressesJazzSettings.Default.ToolTipHelpExit);

            string help_file_name = FileUtil.SubDirectoryFileName(AddressesJazzSettings.Default.FileHelp,
                AddressesJazzSettings.Default.HelpDir, i_exe_dir);

            m_test_protokoll_name = FileUtil.SubDirectoryFileName(AddressesJazzSettings.Default.FileTest,
                                 AddressesJazzSettings.Default.OutputDir, i_exe_dir);

            string help_file_resources = Properties.Resources.JAZZ_live_AARAU_Adressen;

            FileUtil.CreateFileFromResourcesStringIfMissing(help_file_name, help_file_resources);

            this.m_rich_text_box_help.LoadFile(help_file_name, RichTextBoxStreamType.RichText);
        }

        /// <summary>Exit from help dialog</summary>
        private void m_button_help_exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>Open Test Protokoll with Notepad</summary>
        private void m_button_test_protokoll_Click(object sender, EventArgs e)
        {
 
            string test_protokoll_resources = Properties.Resources.TestProtokoll;

            FileUtil.CreateFileFromResourcesStringIfMissing(m_test_protokoll_name, test_protokoll_resources);

            System.Diagnostics.Process.Start("notepad.exe", m_test_protokoll_name);
        }

        /// <summary>Open web page with the default browser</summary>
        private void m_button_program_documentation_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process myProcess = new System.Diagnostics.Process();

            try
            {
                // true is the default, but it is important not to set it to false
                myProcess.StartInfo.UseShellExecute = true;
                myProcess.StartInfo.FileName = AddressesJazzSettings.Default.ProgramDocumentation;
                myProcess.Start();
            }
            catch (Exception e_wiki)
            {
                MessageBox.Show(e_wiki.Message);
            }
        }
    }
}
