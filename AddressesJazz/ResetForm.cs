using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace AddressesJazz
{
    /// <summary>Reset to an old addresses backup file</summary>
    public partial class ResetForm : Form
    {
        #region Member variables

        private JazzForm m_jazz_form = null;
        private JazzMain m_main = null;

        private string[] m_backup_files;

        #endregion // Member variables

        #region Constructor

        /// <summary>Constructor that gets backup files from the server</summary>
        public ResetForm(JazzForm i_jazz_form, JazzMain i_main)
        {
            InitializeComponent();

            m_jazz_form = i_jazz_form;

            m_main = i_main;

            this.m_combo_box_backups_label.Text = AddressesJazzSettings.Default.Caption_BackupFileSelect;

            _SetComboBoxBackups();

        } // Constructor

        #endregion // Constructor

        #region Set controls

        /// <summary>Set the combobox with backup file names</summary>
        private void _SetComboBoxBackups()
        {
            string error_message= "";

            if (!m_main.RemoveTemporaryUsedBackupFiles(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            if (!Reset.DownloadBackupFiles(out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            if (!Reset.GetBackupFiles(out m_backup_files, out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            m_combo_box_backups.Items.Clear();
            foreach (string file_name in m_backup_files)
            {
                string file_name_without_path = Path.GetFileName(file_name);
                m_combo_box_backups.Items.Add(file_name_without_path);
            }

            m_combo_box_backups.Text = "";

        } // _SetComboBoxBackups

        #endregion // Set controls

        #region Event functions

        /// <summary>The user selected a backup file</summary>
        private void m_combo_box_backups_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (m_combo_box_backups.Text.Trim() == "")
                return;

            int file_name_index = m_combo_box_backups.SelectedIndex;
            string file_name_with_path = m_backup_files[file_name_index];

            string error_message = "";
            if (!Reset.ResetWithBackupFile(m_main, file_name_with_path, out error_message))
            {
                MessageBox.Show(error_message);
                return;
            }

            m_jazz_form.SetTable(m_main.GetTable()); // TODO Not so nice ...

            int row_index = 1;
            m_jazz_form.SetControlsTexts(row_index);

            this.Close();

        } //  m_combo_box_backups_SelectedIndexChanged

        #endregion // Event functions

    } // ResetForm

} // namespace
