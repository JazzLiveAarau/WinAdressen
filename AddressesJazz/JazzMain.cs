using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using AdressesUtility;
using Ftp;
using ExcelUtil;
using System.Collections;
using System.Windows.Forms;

namespace AddressesJazz
{
    /// <summary>Main class for the application Jazz Adressen. that executes most of the commands in this application
    /// <para>This class executes most of the commands coming from the controls of the JazzForm form.</para>
    /// <para>The other command execution classes are Output and Reset. In the GUI JazzForm class should nothing be executed.</para>
    /// <para>Typical commands are: Add or delete a person in the address database, sort records, output lists, etc.</para>
    /// <para>At start of the application the current address list (a CSV file) will be downloaded with FTP.</para>
    /// <para>The downloaded address list defines an instance of the class Table that holds all data of the list.</para>
    /// <para>Changes of the address list can only be done if the address list has been checked out.</para>
    /// <para>Checkout means that the user will be registered in an checkin/checkout file also down- and uploaded with FTP</para>
    /// <para>Only one person at a time can checkout the address list.</para>
    /// </summary>
    public class JazzMain
    {
        #region Member parameters

        /// <summary>FTP password for the download and upload</summary>
        static public string m_ftp_password = "TODO";

        /// <summary>Get FTP password</summary>
        public string FtpPassword
        { get { return m_ftp_password; } }

        /// <summary>Addresses table. The table corresponds to the input/otput CSV file.</summary>
        private Table m_table_addresses = null;

        /// <summary>Path to the exe directory. Used to get the paths to the application subdirectories (Excel, Output, etc)</summary>
        static public string m_exe_directory = System.Windows.Forms.Application.StartupPath;

        /// <summary>The main form for this application. Commands are coming from controls (edit fields, buttons, ..) of this form.</summary>
        private JazzForm m_main_form = null;

        /// <summary>Encoding used for read and write for files</summary>
        private Encoding m_file_encoding = Encoding.UTF8;

        /// <summary>Flag telling if the addresses list has been checked out.</summary>
        private bool m_addresses_checked_out = false;

        /// <summary>Get or set the flag telling if the addresses list has been checked out</summary>
        public bool AddressesCheckedOut
        { get { return m_addresses_checked_out; } set { m_addresses_checked_out = value; } }

        #endregion // Member parameters

        #region Constructor

        /// <summary>Constructor 
        /// <para>The config file will be created by a function in class AddressesJazzSettings if not existing in the exe directory. </para>
        /// <para>After the installation of a new version of the application Jazz Adressen the file should (must) be created.</para>
        /// </summary>
        /// <param name="i_main_form">The main form for application Jazz Adressen</param>
        public JazzMain(JazzForm i_main_form)
        {
            this.m_main_form = i_main_form;

            string config_file = AdressesUtility.FileUtil.ConfigFileName(AddressesJazzSettings.Default.ConfigRootElement, JazzMain.m_exe_directory);
            if (!File.Exists(config_file))
            {
                AddressesJazzSettings.Default.Save();
            }
        } // Constructor

        #endregion // Constructor

        #region Create and upload the address CSV file

        /// <summary>Create and upload the address CSV file
        /// <para>Upload of the address file is only allowed if addresses has been checked out.</para>
        /// <para>1. The function checks if there is an Internet connection.</para>
        /// <para>2. The checkin/checkout log file is downloaded with FTP. Return with error message if not checked out.</para>
        /// <para>3. The address CSV file and a backup copy is created by the function TableToCsv in class FromTable.</para>
        /// <para>4. The CSV files are uploaded by the function UploadFile in the class Ftp.UpLoad.</para>
        /// <para>5. A checkin line is added in the checkin/checkout log file. This file is also upoaded with UploadFile.</para>
        /// <para>6. The flag m_addresses_checked_out is set to false.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool CreateAndUploadFiles(out string o_error)
        {
            o_error = "";

            if (false == this.AddressesCheckedOut)
            {
                o_error = AddressesJazzSettings.Default.ErrMsgAddressesNotCheckedOut;
                return false;
            }


            string local_file_name = "";
            string server_path_file = "";

            if (!_DownloadCheckOutInLogFile(out server_path_file, out local_file_name, out o_error))
            {
                return false;
            }

            if (!_AddressesAreCheckedOut(local_file_name, out o_error))
            {
                return false;
            }

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            string local_file_address = _GetAddressesFileName();

            if (AddressesJazzSettings.Default.BetaVersion)
            {
                local_file_address = _GetAddressesFileNameBeta();
            }

            string backup_file_address = _GetBackupAddressesFileName();

            string delimiter = ",";
            if (!FromTable.TableToCsv(m_table_addresses, local_file_address, delimiter, m_file_encoding, out o_error))
                return false;

            if (!FromTable.TableToCsv(m_table_addresses, backup_file_address, delimiter, m_file_encoding, out o_error))
                return false;

            Ftp.UpLoad ftp_upload = new Ftp.UpLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            string server_name_file_csv = AdressesUtility.FileUtil.GetServerFileName(local_file_address);
            string server_file_path_csv = AddressesJazzSettings.Default.AddressesServerDir + server_name_file_csv;

            if (!ftp_upload.UploadFile(server_file_path_csv, local_file_address, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadAddressesFailed;
                return false;
            }

            string backup_dir = AddressesJazzSettings.Default.AddressesBackupsDir;

            string server_name_backup_file_csv = AdressesUtility.FileUtil.GetServerBackupFileName(backup_file_address, backup_dir);
            string server_backup_file_path_csv = AddressesJazzSettings.Default.AddressesServerDir + server_name_backup_file_csv;

            if (!ftp_upload.UploadFile(server_backup_file_path_csv, backup_file_address, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadBackupAddressesFailed;
                return false;
            }

            if (!_AppendLogFileRow(local_file_name, AddressesJazzSettings.Default.CheckInLogFile, out o_error))
                return false;

            if (!ftp_upload.UploadFile(server_path_file, local_file_name, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadLogfile;
                return false;
            }

            string xml_error_msg = "";

            bool b_supporter_xml = CreateAndUploadSupporterXmlFile(m_table_addresses, out xml_error_msg);

            AddressesCheckedOut = false;

            return true;
        } // CreateAndUploadFiles

        #endregion // Create and upload the address CSV file

        #region Create and upload supporter XML file

        /// <summary>
        /// Creates and uploads a supporter XML file to the server
        /// </summary>
        /// <param name="i_table_addresses">Table wit all addresses</param>
        /// <param name="o_error">Error message for failure</param>
        /// <returns>True if the XML file was uploaded</returns>
        public bool CreateAndUploadSupporterXmlFile(Table i_table_addresses, out string o_error)
        {
            o_error = "";

            DateTime current_time = DateTime.Now;

            int current_year = current_time.Year;

            string season_column_name_one = "Beitrag-" + current_year.ToString() + "-" + (current_year + 1).ToString();

            string season_column_name_two = "Beitrag-" + (current_year - 1).ToString() + "-" + current_year.ToString();

            string output_full_file_name_local = GetOutputFileNameNoTimeStamp("Supporter.xml");

            int season_start_year = current_year;

            if (!Output.SupportersAsXml(i_table_addresses, season_column_name_one, output_full_file_name_local, out o_error))
            {
                season_start_year = current_year - 1;

                if (!Output.SupportersAsXml(i_table_addresses, season_column_name_two, output_full_file_name_local, out o_error))
                {
                    return false;
                }
            }

            Ftp.UpLoad ftp_upload = new Ftp.UpLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            string server_name_file_xml = AdressesUtility.FileUtil.GetServerFileName(output_full_file_name_local);

            string server_dir = "/www/QrCode/QrFiles/Season_" 
                + season_start_year.ToString() + "_" + (season_start_year+1).ToString() + "/";

            string ftp_response = "";

            if (!ftp_upload.DoesDirectoryExist(server_dir, out ftp_response))
            {
                string error_dir_str = "";

                if (!ftp_upload.CreateDirectory(server_dir, out error_dir_str))
                {
                    return false;
                }
            }

            string server_file_path_xml = server_dir + server_name_file_xml;

            if (!ftp_upload.UploadFile(server_file_path_xml, output_full_file_name_local, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadAddressesFailed;
                return false;
            }

            return true;

        } // CreateAndUploadSupporterXmlFile

        #endregion // Create and upload supporter XML files

        #region Check in without saving addresses

        /// <summary>Check in without saving addresses
        /// <para>Function for the case that addresses are checked out but the user don't want to upload (save) the file.</para>
        /// <para>The checkin/checkout log file is downloaded with FTP and a checkin line is appended.</para>
        /// <para>The checkin/checkout log file is upoaded with UploadFile in class Ftp.UpLoad.</para>
        /// <para>The flag m_addresses_checked_out is set to false.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool UploadCheckOutInFileButNotAddresses(out string o_error)
        {
            o_error = "";

            if (false == this.AddressesCheckedOut)
            {
                o_error = AddressesJazzSettings.Default.ErrMsgAddressesNotCheckedOut;
                return false;
            }

            string local_file_name = "";
            string server_path_file = "";

            if (!_DownloadCheckOutInLogFile(out server_path_file, out local_file_name, out o_error))
            {
                return false;
            }

            if (!_AddressesAreCheckedOut(local_file_name, out o_error))
            {
                return false;
            }

            Ftp.UpLoad ftp_upload = new Ftp.UpLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            if (!_AppendLogFileRow(local_file_name, AddressesJazzSettings.Default.CheckInLogFile, out o_error))
                return false;

            if (!ftp_upload.UploadFile(server_path_file, local_file_name, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadLogfile;
                return false;
            }

            AddressesCheckedOut = false;

            return true;

        } // UploadCheckOutInFileButNotAddresses

        #endregion // Check in without saving addresses

        #region Download the addresses CSV file

        /// <summary>Download the addresses CSV file from the server with FTP
        /// <para>1. The function checks if there is an Internet connection.</para>
        /// <para>2. The addresses CSV file is downloaded with function DownloadFile in class Ftp.DownLoad.</para>
        /// <para>3. The file is saved in a subdirectory (Excel) to the exe directory.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool DownloadFile(out string o_error)
        {
            o_error = "";

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            if (AddressesJazzSettings.Default.CopyCurrentAddressesForBetaVersion)
            {
                // TODO This should normally be done only once when Beta testing starts
            }

            string server_address_file = AddressesJazzSettings.Default.AddressesFileName;

            string server_address_path_file = AddressesJazzSettings.Default.AddressesServerDir + AddressesJazzSettings.Default.AddressesFileName;

            if (AddressesJazzSettings.Default.BetaVersion)
            {
                server_address_file = AddressesJazzSettings.Default.AddressesFileNameBeta;

                server_address_path_file = AddressesJazzSettings.Default.AddressesServerDir + AddressesJazzSettings.Default.AddressesFileNameBeta;
            }

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            string local_file_with_path = _GetAddressesFileName();

            if (AddressesJazzSettings.Default.BetaVersion)
            {
                local_file_with_path = _GetAddressesFileNameBeta();
            }

            if (!ftp_download.DownloadFile(server_address_path_file, local_file_with_path, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoExcelFileDownload;
                return false;
            }

            return true;

        } // DownloadFile

        #endregion // Download the addresses CSV file

        #region Download installer

        /// <summary>Download an installer for a new version of the Jazz Adressen application from the server with FTP
        /// <para>1. The function checks if there is an Internet connection.</para>
        /// <para>2. The installer is (and possibly other files are) downloaded with function Getfiles in class Ftp.DownLoad.</para>
        /// <para>3. The installer is saved in a subdirectory (NeueVersion) to the exe directory.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool DownloadNewVersion(out string o_error)
        {
            o_error = "";

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            bool b_down_load = true;

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            string server_address_directory = @"/" + AddressesJazzSettings.Default.AdressenServerDir + @"/" +
                                                  AddressesJazzSettings.Default.NewVersionDir + @"/";

            string local_address_directory = FileUtil.SubDirectory(AddressesJazzSettings.Default.NewVersionDir, m_exe_directory);

            if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
            {
                b_down_load = false;
            }

            if (!b_down_load)
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNewVersionDownload;
            }

            return b_down_load;

        } // DownloadNewVersion


        #endregion // Download installer

        #region Checkout and checkin

        /// <summary>Get the error message when the user has checked out addresses and exits the application with cancel.</summary>
        /// <para>1. Message and caption is retrieved from the configuration file (AddressesJazzSettings).</para>
        /// <param name="o_message">Message</param>
        /// <param name="o_caption">Caption for the error message</param>
        /// <param name="o_error">Error description</param>
        public bool GetSaveAddressesMessage(out string o_message, out string o_caption, out string o_error)
        {
            o_error = "";
            o_message = "";
            o_caption = "";

            if (AddressesCheckedOut == false)
            {
                o_error = "GetCheckOutAddressesMessage Programming error: Addresses are not checked out";
                return false;
            }

            o_message = AddressesJazzSettings.Default.MsgShallAddressesBeUploaded;
            o_caption = AddressesJazzSettings.Default.MsgCaptionShallAddressesBeUploaded;

            return true;

        } // GetSaveAddressesMessage

        /// <summary>Check out the addresses file 
        /// <para>1. The checkin/checkout file is downloaded (with function _DownloadCheckOutInLogFile).</para>
        /// <para>2. The fields of the last row in the file is retrieved (with function _GetLastRowCheckInOutFields)</para>
        /// <para>3. Return with error if addresses already are checked out by somebody else.</para>
        /// <para>4. Append logout line with function _AppendLogFileRow</para>
        /// <para>5. Upload the checkin/checkout file with function UploadFile in class Ftp.UpLoad.</para>
        /// <para>6. The flag m_addresses_checked_out is set to true.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool CheckOutAddresses(out string o_error)
        {
            o_error = "";

            string local_file_name = "";
            string server_path_file = "";

            if (!_DownloadCheckOutInLogFile(out server_path_file, out local_file_name, out o_error))
            {
                return false;
            }

            string row_start = "";
            string row_time = "";
            string row_machine = "";
            if (!_GetLastRowCheckInOutFields(local_file_name, out row_start, out row_time, out row_machine, out o_error))
            {
                return false;
            }

            string comp_string_in = AddressesJazzSettings.Default.CheckInLogFile;
            string comp_string_out = AddressesJazzSettings.Default.CheckOutLogFile;

            if (String.Compare(comp_string_in, row_start, false) == 0)
            {
                if (!_AppendLogFileRow(local_file_name, comp_string_out, out o_error))
                    return false;
            }
            else if (String.Compare(comp_string_out, row_start, false) == 0)
            {
                o_error = AddressesJazzSettings.Default.ErrMsgAddressesCheckOutBy + row_machine + "\n" +
                          row_time + "\n\n" + AddressesJazzSettings.Default.MsgAddressesForceCheckOut;
                return false;
            }
            else
            {
                o_error = @"Programming error: Last row start element in login-logout is " + row_start;
                return false;
            }


            Ftp.UpLoad ftp_upload = new Ftp.UpLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            if (!ftp_upload.UploadFile(server_path_file, local_file_name, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadLogfile;
                return false;
            }

            AddressesCheckedOut = true;

            return true;

        } // CheckOutAddresses

        /// <summary>Force a checkout even if addresses are checked out by somebody else, 
        /// <para>i.e. check out the addresses file without checking if is checked in</para>
        /// <para>There should normally be no need to call this function! End of application shall always make a checkin.</para>
        /// <para>1. The checkin/checkout file is downloaded (with function _DownloadCheckOutInLogFile).</para>
        /// <para>2. Append logout line with function _AppendLogFileRow</para>
        /// <para>3. Upload the checkin/checkout file with function UploadFile in class Ftp.UpLoad.</para>
        /// <para>4. The flag m_addresses_checked_out is set to true.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool ForceCheckOutAddresses(out string o_error)
        {
            o_error = "";

            string local_file_name = "";
            string server_path_file = "";

            if (!_DownloadCheckOutInLogFile(out server_path_file, out local_file_name, out o_error))
            {
                return false;
            }

            string comp_string_out = AddressesJazzSettings.Default.CheckOutLogFile;

            if (!_AppendLogFileRow(local_file_name, comp_string_out, out o_error))
                return false;

            Ftp.UpLoad ftp_upload = new Ftp.UpLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            if (!ftp_upload.UploadFile(server_path_file, local_file_name, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgUploadLogfile;
                return false;
            }

            AddressesCheckedOut = true;

            return true;

        } // ForceCheckOutAddresses

        /// <summary>Returns false if addresses not are checked out by this computer
        /// <para>1. The fields of the last row in the file is retrieved (with function _GetLastRowCheckInOutFields)</para>
        /// <para>2. Return with error if computer field name not is equal to this computer.</para>
        /// </summary>
        /// <param name="i_local_file_name">Full local name of the checkin-checkout file</param>
        /// <param name="o_error">Error description</param>
        private bool _AddressesAreCheckedOut(string i_local_file_name, out string o_error)
        {
            o_error = "";

            string row_start = "";
            string row_time = "";
            string row_machine = "";
            if (!_GetLastRowCheckInOutFields(i_local_file_name, out row_start, out row_time, out row_machine, out o_error))
            {
                return false;
            }

            string comp_string_in = AddressesJazzSettings.Default.CheckInLogFile;
            string comp_string_out = AddressesJazzSettings.Default.CheckOutLogFile;

            if (String.Compare(comp_string_in, row_start, false) == 0)
            {
                o_error = @"_AddressesAreCheckedOut Programming error: Addresses are not checked out";
                return false;
            }
            else if (String.Compare(comp_string_out, row_start, false) == 0)
            {
                string machine = System.Environment.MachineName;

                if (String.Compare(machine, row_machine, false) != 0)
                {
                    o_error = @"_AddressesAreCheckedOut Programming error: Addresses are checked out by " +
                             row_machine + @" and not you (" + machine + @")";
                    return false;
                }
            }
            else
            {
                o_error = @"_AddressesAreCheckedOut Programming error: Last row start element in login-logout is " + row_start;
                return false;
            }

            return true;

        } // _AddressesAreCheckedOut

        /// <summary>Download the checkin/checkout file from the server with FTP
        /// <para>1. The function checks if there is an Internet connection.</para>
        /// <para>2. The checkin/checkout file is downloaded with function DownloadFile in class Ftp.DownLoad.</para>
        /// </summary>
        /// <param name="o_server_path_file">Server checkin-checkout path-file name</param>
        /// <param name="o_local_file_name">Full name of checkin-checkout file</param>
        /// <param name="o_error">Error description</param>
        public bool _DownloadCheckOutInLogFile(out string o_server_path_file, out string o_local_file_name, out string o_error)
        {
            o_local_file_name = _GetCheckInOutLogFileName();

            o_server_path_file = AddressesJazzSettings.Default.AddressesServerDir + Path.GetFileName(o_local_file_name);

            o_error = "";

            Ftp.DownLoad ftp_download = new Ftp.DownLoad(AddressesJazzSettings.Default.FtpHost, AddressesJazzSettings.Default.FtpUser, m_ftp_password);

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            if (!ftp_download.DownloadFile(o_server_path_file, o_local_file_name, out o_error))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoCheckInOutLogFileDownload;
                return false;
            }

            return true;

        } // _DownloadCheckOutInLogFile


        /// <summary>Append row to login-logout file
        /// <para>Time and machine (computer) is added to the input string. </para>
        /// <para>The line is appended to the file.</para>
        /// </summary>
        /// <param name="i_local_file_name">Full input file name</param>
        /// <param name="i_start_append_row">Start string for the append row (login or logout)</param>
        /// <param name="o_error">Error description</param>
        private bool _AppendLogFileRow(string i_local_file_name, string i_start_append_row, out string o_error)
        {
            o_error = "";

            string append_row = "\n" + i_start_append_row + @" " + TimeUtil.YearMonthDayHourMinSec() + 
                                @" " + System.Environment.MachineName;

            try
            {
                using (StreamWriter writer = File.AppendText(i_local_file_name))
                {
                    writer.Write(append_row);
                }
            }

            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            return true;

        } // _AppendLogFileRow

        /// <summary>Get the fields of the last row of the checkin-checkout file 
        /// <para>1. Get last row (call of function _GetLastRow).</para>
        /// <para>2. Retrieve the field values. The start of each row is defined by configuration</para>
        /// <para>   file strings CheckInLogFile and CheckOutLogFile</para>
        /// </summary>
        /// <param name="i_local_file_name">Full input file name</param>
        /// <param name="o_start">Start field of the row</param>
        /// <param name="o_time">Time field of the row</param>
        /// <param name="o_machine">Machine name field of the row</param>
        /// <param name="o_error">Error description</param>
        private bool _GetLastRowCheckInOutFields(string i_local_file_name, out string o_start, out string o_time, out string o_machine, out string o_error)
        {
            o_start = "";
            o_time = "";
            o_machine = "";
            o_error = "";

            string last_row = "";

            if (!_GetLastRow(i_local_file_name, out last_row, out o_error))
            {
                return false;
            }

            string comp_string_in = AddressesJazzSettings.Default.CheckInLogFile;
            string start_string_in = last_row.Substring(0, comp_string_in.Length);

            string comp_string_out = AddressesJazzSettings.Default.CheckOutLogFile;
            string start_string_out = last_row.Substring(0, comp_string_out.Length);

            if (String.Compare(comp_string_in, start_string_in, false) == 0)
            {
                o_start = comp_string_in;
            }
            else if (String.Compare(comp_string_out, start_string_out, false) == 0)
            {
                o_start = comp_string_out;
            }
            else
            {
                o_error = @"Programming error: Last row in login-logout is " + last_row;
                return false;
            }

            string current_field = "";

            for (int i_char = o_start.Length + 1; i_char < last_row.Length; i_char++)
            {
                string current_char = last_row.Substring(i_char, 1);

                if (current_char.CompareTo(" ") == 0)
                {
                    if (o_time.CompareTo("") == 0)
                    {
                        o_time = current_field;
                        current_field = "";
                    }
                }
                else
                {
                    current_field = current_field + current_char;
                }
            }

            o_machine = current_field;

            return true;

        } // _GetLastRowCheckInOutFields

        /// <summary>Get last row of the input file </summary>
        /// <para>File is opened and all rows are read. The last (not empy) row is returned.</para>
        /// <param name="i_local_file_name">Full input file name</param>
        /// <param name="o_last_row">Last (non-empty) row of the file</param>
        /// <param name="o_error">Error description</param>
        private bool _GetLastRow(string i_local_file_name, out string o_last_row, out string o_error)
        {
            o_last_row = "";
            o_error = "";

            if (!File.Exists(i_local_file_name))
            {
                o_error = @"File: " + i_local_file_name + @" does not exist. Programming error";
                return false;
            }

            try
            {
                using (FileStream file_stream = new FileStream(i_local_file_name, FileMode.Open, FileAccess.Read, FileShare.Read))
                // Without System.Text.Encoding.UTF8 there are problems with ä ö ü. With Encoding.Default it worked in some computers
                // Alternatives Encoding.Default, Encoding.UTF8, Encoding.Unicode, Encoding.UTF32, Encoding.UTF7
                // using (StreamReader stream_reader = new StreamReader(file_stream, System.Text.Encoding.Default))
                using (StreamReader stream_reader = new StreamReader(file_stream))
                {
                    while (stream_reader.Peek() >= 0)
                    {
                        string current_row = stream_reader.ReadLine();

                        if (current_row.Trim() == "")
                        {
                            // A line with only spaces.
                            break;
                        }
                        else
                        {
                            o_last_row = current_row;
                        }

                    } // while
                }
            }


            catch (FileNotFoundException) { o_error = "File not found"; return false; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return false; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return false; }
            catch (InvalidCastException) { o_error = "invalid cast"; return false; }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            return true;

        } // _GetLastRow

        #endregion // Checkout and checkin

        #region File functions and names

        /// <summary>Returns false if the local address file is missing</summary>
        /// <param name="o_error">Error description</param>
        public bool LocalFileExists(out string o_error)
        {
            o_error = "";

            if (!File.Exists(_GetAddressesFileName()))
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoLocalExcelFile + _GetAddressesFileName() + ". " + AddressesJazzSettings.Default.MsgExitApplication;
                return false;
            }
            else
            {
                return true;
            }

        } // LocalFileExists

        /// <summary>Returns the full file name for an output file
        /// <para>1. Output directory is retrieved from the configuration file (AddressesJazzSettings).</para>
        /// <para>2. A help function in class AdressesUtility.FileUtil is called to construct the name.</para>
        /// </summary>
        /// <param name="i_file_name">File name with extension without path</param>
        public string GetOutputFileName(string i_file_name)
        {
            string addresses_directory = AddressesJazzSettings.Default.OutputDir;
            string file_name_date = Path.GetFileNameWithoutExtension(i_file_name) + TimeUtil.YearMonthDay() + Path.GetExtension(i_file_name);
            string local_address_file = AdressesUtility.FileUtil.AddressesFileName(file_name_date, addresses_directory, JazzMain.m_exe_directory);
   
            return local_address_file;

        } // GetOutputFileName

        /// <summary>Returns the full file name for an output file without a time stamp
        /// <para>1. Output directory is retrieved from the configuration file (AddressesJazzSettings).</para>
        /// <para>2. A help function in class AdressesUtility.FileUtil is called to construct the name.</para>
        /// </summary>
        /// <param name="i_file_name">File name with extension without path</param>
        public string GetOutputFileNameNoTimeStamp(string i_file_name)
        {
            string addresses_directory = AddressesJazzSettings.Default.OutputDir;
           string local_address_file = AdressesUtility.FileUtil.AddressesFileName(i_file_name, addresses_directory, JazzMain.m_exe_directory);

            return local_address_file;

        } // GetOutputFileNameNoTimeStamp


        /// <summary>Returns the full local file name for the addresses file
        /// <para>1. File name and subdirectory name from the configuration file (AddressesJazzSettings).</para>
        /// <para>2. Create the combined full name. Call of function AdressesUtility.FileUtil.AddressesFileName.</para>
        /// </summary>
        private string _GetAddressesFileName()
        {
            string addresses_file_name = AddressesJazzSettings.Default.AddressesFileName;
            string addresses_directory = AddressesJazzSettings.Default.AddressesDir;
            string local_address_file = AdressesUtility.FileUtil.AddressesFileName(addresses_file_name, addresses_directory, JazzMain.m_exe_directory);

            return local_address_file;

        } // _GetAddressesFileName

        /// <summary>Returns the full file local name for the Beta version addresses file
        /// <para>1. File name and subdirectory name from the configuration file (AddressesJazzSettings).</para>
        /// <para>2. Create the combined full name. Call of function AdressesUtility.FileUtil.AddressesFileName.</para>
        /// </summary>
        private string _GetAddressesFileNameBeta()
        {
            string addresses_file_name = AddressesJazzSettings.Default.AddressesFileNameBeta;
            string addresses_directory = AddressesJazzSettings.Default.AddressesDir;
            string local_address_file = AdressesUtility.FileUtil.AddressesFileName(addresses_file_name, addresses_directory, JazzMain.m_exe_directory);

            return local_address_file;

        } // _GetAddressesFileNameBeta

        /// <summary>Returns the file name for the FTP addresses file
        /// <para>1. File name from the configuration file (AddressesJazzSettings).</para>
        /// </summary>
        private string _GetFtpAddressesFileName()
        {
            return AddressesJazzSettings.Default.AddressesFileName;

        } // _GetFtpAddressesFileName

        /// <summary>Returns the file name for the checkin-checkout logfile
        /// <para>1. File name and subdirectory name from the configuration file (AddressesJazzSettings).</para>
        /// <para>2. Create the combined full name. Call of function AdressesUtility.FileUtil.AddressesFileName.</para>
        /// </summary>
        private string _GetCheckInOutLogFileName()
        {
            string addresses_directory = AddressesJazzSettings.Default.AddressesDir;
            string local_address_file = AdressesUtility.FileUtil.AddressesFileName(AddressesJazzSettings.Default.CheckInOutLogFileName, addresses_directory, JazzMain.m_exe_directory);

            return local_address_file;

        } // _GetCheckInOutLogFileName

        /// <summary>Returns the full file name for the backup addresses file
        /// <para>1. File name and subdirectory name from the configuration file (AddressesJazzSettings).</para>
        /// <para>   (Another file name for the beta version of the application).</para>
        /// <para>2. Create the combined full name. Call of function AdressesUtility.FileUtil.BackupAddressesFileName.</para>
        /// </summary>
        private string _GetBackupAddressesFileName()
        {
            string addresses_file_name = AddressesJazzSettings.Default.AddressesFileName;
            if (AddressesJazzSettings.Default.BetaVersion)
            {
                addresses_file_name = AddressesJazzSettings.Default.AddressesFileNameBeta;
            }
            string addresses_directory = AddressesJazzSettings.Default.AddressesBackupsDir;
            string local_time_address_file = AdressesUtility.FileUtil.BackupAddressesFileName(addresses_file_name, addresses_directory, JazzMain.m_exe_directory);

            return local_time_address_file;

        } // _GetBackupAddressesFileName

        #endregion // File functions and names

        #region Create table

        /// <summary>Create addresses table from the addresses list
        /// <para>1. Create and set the row header. Call of _SetRowHeader.</para>
        /// <para>2. An instance of class Table is created (member variable m_table_addresses)</para>
        /// <para>3. Get the local name of the csv file. Call of _GetAddressesFileName.</para>
        /// <para>4. Create the table from the (from the server downloaded) csv file. Call of ToTable.CsvToTable.</para>
        /// <para>5. For a new current season add a supporter column. Call of _AddSupporterColumnForNewCurrentSeason</para>
        /// <para>6. Set header for the table (although not yet used in the application). Call of Table.SetRowHeader.</para>
        /// <para></para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool CreateAddressesTable(out string o_error)
        {
            o_error = "";

            RowHeader row_header = null;
            _SetRowHeader(out row_header);

            m_table_addresses = new Table("Addresses");

            string local_file_with_path = _GetAddressesFileName();
            if (AddressesJazzSettings.Default.BetaVersion)
            {
                local_file_with_path = _GetAddressesFileNameBeta();
            }

            // Todo UTF 8 should be right if (!ToTable.CsvToTable(_GetAddressesFileName(), ref m_table_addresses, ref m_file_encoding, out o_error))
            if (!ToTable.CsvToTable(local_file_with_path, ref m_table_addresses, out o_error))
            {
                o_error = "JazzMain.CreateAddressesTable ToTable.CsvToTable failed " + o_error;
                return false;
            }
                
            if (!_AddSupporterColumnForNewCurrentSeason(row_header, out o_error))
            {
                o_error = "JazzMain.CreateAddressesTable _AddSupporterColumnForNewCurrentSeason failed " + o_error;
                return false;
            }

            m_table_addresses.SetRowHeader(row_header);

            return true;

        } // CreateAddressesTable

        /// <summary>For a new current season a supporter column will be added
        /// <para>The table m_table_addresses is the input and output table.</para>
        /// <para>The input header row determines if a supporter column shall be added</para>
        /// <para>1. Return without doing anything if the number of columns of the header row and the table is equal.</para>
        /// <para>2. Get the field header name from the input header row. Call of FileHeader.GetFieldHeader.Name</para>
        /// <para>3. Create string array with values for the column that shall be appended. Call of _SupportColumnInitialValuesAsStrings.</para>
        /// <para>4. Create the column to append. Call of ExcelUtil.TableTools.CreateColumn.</para>
        /// <para>5. Append the column. Call of ExcelUtil.TableTools.InsertColumn.</para>
        /// <para>6. Create the csv file with the added column. Call of FromTable.TableToCsv</para>
        /// <para>It is important that this file is created. Functions may create table from the file.</para>
        /// <para>Please note that the csv file with the appended column will be saved on the server with checkout->checkin (and not by this function).</para>
        /// </summary>
        /// <param name="i_row_header">Header row</param>
        /// <param name="o_error">Error description</param>
        private bool _AddSupporterColumnForNewCurrentSeason(RowHeader i_row_header, out string o_error)
        {
            o_error = "";

            int n_columns_header = i_row_header.NumberColumns;

            int n_columns_table = m_table_addresses.NumberColumns;

            if (n_columns_header == n_columns_table)
            {
                // Do not append a column
                return true;
            }

            if (n_columns_header != n_columns_table + 1)
            {
                o_error = "JazzMain._AddSupporterColumnForNewCurrentSeason n_columns_header = " + n_columns_header.ToString() +
                            " != n_columns_table + 1 = " + (n_columns_table + 1).ToString();
                return false;
            }

            string error_field_header = "";
            FieldHeader field_header = i_row_header.GetFieldHeader(n_columns_header - 1, out error_field_header);
            if (error_field_header.Length > 0)
            {
                o_error = "JazzMain._AddSupporterColumnForNewCurrentSeason error_field_header = " + error_field_header;
                return false;
            }

            int n_rows = m_table_addresses.NumberRows;
            string[] fields_as_strings = _SupportColumnInitialValuesAsStrings(n_rows, field_header.Name);

            Column append_column;
            if (!ExcelUtil.TableTools.CreateColumn(fields_as_strings, out append_column, out o_error))
            {
                o_error = "JazzMain._AddSupporterColumnForNewCurrentSeason ExcelUtil.TableTools.CreateColumn failed " + o_error;
                return false;
            }

            int index_column = m_table_addresses.NumberColumns;
            if (!ExcelUtil.TableTools.InsertColumn(ref m_table_addresses, index_column, append_column, out o_error))
            {
                o_error = "JazzMain._AddSupporterColumnForNewCurrentSeason ExcelUtil.TableTools.InsertColumn failed " + o_error;
                return false;
            }

            string local_file_address = _GetAddressesFileName();
            string delimiter = ",";
            if (!FromTable.TableToCsv(m_table_addresses, local_file_address, delimiter, m_file_encoding, out o_error))
            {
                o_error = "JazzMain.CreateAddressesTable FromTable.TableToCsv failed " + o_error;
                return false;
            }

            return true;

        } // _AddSupporterColumnForNewCurrentSeason

        /// <summary>Create array of default values for a support column
        /// <para>The first element of the array (column) will get the value i_field_header.</para>
        /// <para>All the other values of the output array (the column) will be '0'</para>
        /// </summary>
        /// <param name="i_number_rows">Number of rows of the table</param>
        /// <param name="i_field_header">The header name of column</param>
        private string[] _SupportColumnInitialValuesAsStrings(int i_number_rows, string i_field_header)
        {
            string[] fields_as_strings;
            ArrayList array_list_fields = new ArrayList();

            for (int i_row = 0; i_row < i_number_rows; i_row++)
            {
                string field_value = "";
                if (0 == i_row)
                {
                    field_value = i_field_header;
                }
                else
                {
                    field_value = "0";
                }

                array_list_fields.Add(field_value);
            }

            fields_as_strings = (string[])array_list_fields.ToArray(typeof(string));

            return fields_as_strings;

        } // _SupportColumnInitialValuesAsStrings

        /// <summary>Returns the table (m_table_addresses) with addresses</summary>
        public Table GetTable()
        {
            return m_table_addresses;

        } // GetTable

        /// <summary>Convert field type string to field type enum</summary>
        private ExcelUtil.FieldType _StringToEnum(string i_str_type)
        {
            ExcelUtil.FieldType ret_type = ExcelUtil.FieldType.UNDEFINED;

            if (i_str_type == "string")
            {
                ret_type = ExcelUtil.FieldType.STRING;
            }
            else if (i_str_type == "float")
            {
                ret_type = ExcelUtil.FieldType.FLOAT;
            }
            else if (i_str_type == "integer")
            {
                ret_type = ExcelUtil.FieldType.INTEGER;
            }
            else if (i_str_type == "boolean")
            {
                ret_type = ExcelUtil.FieldType.BOOLEAN;
            }

            return ret_type;

        } // _StringToEnum

        /// <summary>Create and set row header for the address table
        /// <para>The data for the first eleven (11) header records are retrieved from the config XML file object (AddressesJazzSettings)</para>
        /// <para>The next header records are supporter records. Such a record holds the amount of money that a person has paid in order to become a supporter.</para>
        /// <para>There is one record for each season starting with 2009-2010. The last header record is for the current season.</para>
        /// <para>Please note that the existing csv file may not have records (a column) corresponding to the last supporter header record.</para>
        /// <para>This occurs when a new season becomes the current season. Function Season.GetCurrentSeasonStartYear defines the current season.</para>
        /// <para>Data for the supporter header records are retrieved with Season functions RecordNameSupporter, RecordTypeSupporter and RecordHelpSupporter.</para>
        /// <para></para>
        /// </summary>
        /// <param name="o_row_header">Created row header</param>
        private void _SetRowHeader(out RowHeader o_row_header)
        {
            o_row_header = new RowHeader();

            FieldHeader field_header_01 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_01);
            field_header_01.Caption = AddressesJazzSettings.Default.Caption_Record_01;
            field_header_01.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_01);
            field_header_01.Help = AddressesJazzSettings.Default.Help_Record_01;
            o_row_header.AddFieldHeader(field_header_01);

            FieldHeader field_header_02 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_02);
            field_header_02.Caption = AddressesJazzSettings.Default.Caption_Record_02;
            field_header_02.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_02);
            field_header_02.Help = AddressesJazzSettings.Default.Help_Record_02;
            o_row_header.AddFieldHeader(field_header_02);

            FieldHeader field_header_03 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_03);
            field_header_03.Caption = AddressesJazzSettings.Default.Caption_Record_03;
            field_header_03.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_03);
            field_header_03.Help = AddressesJazzSettings.Default.Help_Record_03;
            o_row_header.AddFieldHeader(field_header_03);

            FieldHeader field_header_04 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_04);
            field_header_04.Caption = AddressesJazzSettings.Default.Caption_Record_04;
            field_header_04.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_04);
            field_header_04.Help = AddressesJazzSettings.Default.Help_Record_04;
            o_row_header.AddFieldHeader(field_header_04);

            FieldHeader field_header_05 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_05);
            field_header_05.Caption = AddressesJazzSettings.Default.Caption_Record_05;
            field_header_05.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_05);
            field_header_05.Help = AddressesJazzSettings.Default.Help_Record_05;
            o_row_header.AddFieldHeader(field_header_05);

            FieldHeader field_header_06 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_06);
            field_header_06.Caption = AddressesJazzSettings.Default.Caption_Record_06;
            field_header_06.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_06);
            field_header_06.Help = AddressesJazzSettings.Default.Help_Record_06;
            o_row_header.AddFieldHeader(field_header_06);

            FieldHeader field_header_07 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_07);
            field_header_07.Caption = AddressesJazzSettings.Default.Caption_Record_07;
            field_header_07.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_07);
            field_header_07.Help = AddressesJazzSettings.Default.Help_Record_07;
            o_row_header.AddFieldHeader(field_header_07);

            FieldHeader field_header_08 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_08);
            field_header_08.Caption = AddressesJazzSettings.Default.Caption_Record_08;
            field_header_08.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_08);
            field_header_08.Help = AddressesJazzSettings.Default.Help_Record_08;
            o_row_header.AddFieldHeader(field_header_08);

            FieldHeader field_header_09 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_09);
            field_header_09.Caption = AddressesJazzSettings.Default.Caption_Record_09;
            field_header_09.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_09);
            field_header_09.Help = AddressesJazzSettings.Default.Help_Record_09;
            o_row_header.AddFieldHeader(field_header_09);

            FieldHeader field_header_10 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_10);
            field_header_10.Caption = AddressesJazzSettings.Default.Caption_Record_10;
            field_header_10.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_10);
            field_header_10.Help = AddressesJazzSettings.Default.Help_Record_10;
            o_row_header.AddFieldHeader(field_header_10);

            FieldHeader field_header_11 = new FieldHeader(AddressesJazzSettings.Default.Name_Record_11);
            field_header_11.Caption = AddressesJazzSettings.Default.Caption_Record_11;
            field_header_11.Type = _StringToEnum(AddressesJazzSettings.Default.Type_Record_11);
            field_header_11.Help = AddressesJazzSettings.Default.Help_Record_11;
            o_row_header.AddFieldHeader(field_header_11);

            int supporter_season_start_year = Season.GetSupporterSeasonStartYear();
            int current_season_start_year = Season.GetCurrentSeasonStartYear();

            for (int i_season= supporter_season_start_year; i_season<= current_season_start_year; i_season++)
            {
                string supporter_record_name = Season.RecordNameSupporter(i_season);
                FieldHeader file_header_supporter = new FieldHeader(supporter_record_name);
                file_header_supporter.Type = _StringToEnum(Season.RecordTypeSupporter());
                file_header_supporter.Caption = Season.SeasonString(i_season);
                file_header_supporter.Help = Season.RecordHelpSupporter(i_season);
                o_row_header.AddFieldHeader(file_header_supporter);
            }

        } // _SetRowHeader

        /// <summary>Replace the existing addresses table with a new table defined by an input (backup) file
        /// <para>For the case that the address CSV file is corrupted a backup file can be used.</para>
        /// <para>1. Create an instance of Table</para>
        /// <para>2. Set (backup) data with function CsvToTable in class ToTable</para>
        /// <para>3. Set m_table_addresses to the new Table instance</para>
        /// </summary>
        /// <param name="i_local_file_with_path">Input CSV file with full path</param>
        /// <param name="o_error">Error description</param>
        public bool ReplaceAddressesTable(string i_local_file_with_path, out string o_error)
        {
            o_error = "";

            RowHeader row_header = null;
            _SetRowHeader(out row_header);

            Table replace_table_addresses = new Table("Addresses", row_header);

            if (!ToTable.CsvToTable(i_local_file_with_path, ref replace_table_addresses, out o_error))
                return false;

            // TODO Should delete instance of m_table_addresses .....
            m_table_addresses = replace_table_addresses;

            return true;

        } // ReplaceAddressesTable

        #endregion // Create table

        #region Search

        /// <summary>Returns a string array with found records and a corresponding array with record indices
        /// <para>The function GetRows in class TableSearch is used to get the records with the given search string.</para>
        /// <para>Search is made in fields first name and family name.</para>
        /// </summary>
        /// <param name="i_search_string">Search string</param>
        /// <param name="o_display_strings">Display strings</param>
        /// <param name="o_record_indices">Record indices</param>
        /// <param name="o_error">Error description</param>
        public bool Search(string i_search_string, out string[] o_display_strings, out int[] o_record_indices, out string o_error)
        {
            ArrayList array_list_display_strings = new ArrayList();
            o_display_strings = (string[])array_list_display_strings.ToArray(typeof(string));

            ArrayList array_list_record_indices = new ArrayList();
            o_record_indices = (int[])array_list_record_indices.ToArray(typeof(int));

            o_error = "";

            string search_string_trimmed = i_search_string.Trim();

            if (search_string_trimmed == "") return true;

            string[] column_headers = new string[3];
            column_headers[0] = AddressesJazzSettings.Default.Name_Record_01;
            column_headers[1] = AddressesJazzSettings.Default.Name_Record_02;
            column_headers[2] = AddressesJazzSettings.Default.Name_Record_07;

            int[] row_indices;
            if (!TableSearch.GetRows(this.m_table_addresses, column_headers, search_string_trimmed, out row_indices, out o_error)) return false;

            o_record_indices = row_indices;

            bool b_error = false;

            Row first_row = this.m_table_addresses.GetRow(0, out o_error);
            if (o_error != "")
            {
                b_error = true;
            }

            for (int i_found = 0; i_found < row_indices.Length && !b_error; i_found++)
            {
                Row current_row = this.m_table_addresses.GetRow(row_indices[i_found], out o_error);
                if (o_error != "")
                {
                    b_error = true;
                    break;
                }

                string display_string = "";

                for (int i_column = 0; i_column < column_headers.Length; i_column++)
                {
                    string column_header = column_headers[i_column];
                    int column_index = Table.GetColumnIndex(first_row, column_header);

                    if (column_index < 0)
                    {
                        o_error = "m_button_search_Click There is no column with header " + column_header;
                        b_error = true;
                        break;
                    }

                    Field current_field = current_row.GetField(column_index, out o_error);
                    if (o_error != "")
                    {
                        b_error = true;
                        break;
                    }

                    if (i_column == column_headers.Length - 1)
                    {
                        display_string = display_string + current_field.FieldValue;
                    }
                    else
                    {
                        display_string = display_string + current_field.FieldValue + ",  ";
                    }
                } // i_column

                array_list_display_strings.Add(display_string);

            } // i_found

            if (b_error)
                return false;

            o_display_strings = (string[])array_list_display_strings.ToArray(typeof(string));

            return true;

        } // Search

        #endregion // Search

        #region Sort table

        /// <summary>Sorts the table
        /// <para>1. Function SortField in class TableSort is called.</para>
        /// </summary>
        /// <param name="i_column_header">Column header string</param>
        /// <param name="o_error">Error message if function fails</param>
        /// <returns>False, if function fails</returns>
        public bool Sort(string i_column_header, out string o_error)
        {
            o_error = "";

            if (!TableSort.SortField(ref this.m_table_addresses, i_column_header, out o_error)) return false;

            return true;

        } // Sort

        #endregion // Sort table

        #region Add and remove table rows

        /// <summary>Append empty row to the table</summary>
        /// <para>1. Function Table.AddEmptyRow is called.</para>
        /// <param name="o_row_index">Index for the added row</param>
        /// <param name="o_error">Error description</param>
        public bool AppendEmptyRow(out int o_row_index, out string o_error)
        {
            o_row_index = -1;
            o_error = "";

            if (!m_table_addresses.AddEmptyRow(out o_row_index, out o_error)) return false;

            return true;

        } // AppendEmptyRow

        /// <summary>Set support for all seasons to zero (0) in a given row
        /// <para>The support is not allowed to be empty</para>
        /// <para>1. Functions Table.GetRow, Table.GetField, Field.FieldValue and Row.SetFieldValue are called.</para>
        /// </summary>
        /// <param name="i_row_index">Index for the input row</param>
        /// <param name="o_error">Error description</param>
        public bool SetSupportToZeroForRow(int i_row_index, out string o_error)
        {
            o_error = "";

            Row first_row = m_table_addresses.GetRow(0, out o_error);
            Row current_row = m_table_addresses.GetRow(i_row_index, out o_error);

            for (int col_index = 0; col_index < first_row.NumberColumns; col_index++)
            {
                Field field_first_row = first_row.GetField(col_index, out o_error);
                string field_value = field_first_row.FieldValue;

                if (field_value.Contains("Beitrag"))
                {
                    current_row.SetFieldValue(col_index, "0", out o_error);
                }               
            }

            return true;

        } // SetSupportToZeroForRow

        /// <summary>Remove row in the table
        /// <para>1. Function Table.RemoveRow is called.</para>
        /// </summary>
        /// <param name="i_row_index">Index for that shall be removed</param>
        /// <param name="o_error">Error description</param>
        public bool RemoveRow(int i_row_index, out string o_error)
        {
            o_error = "";

            if (!m_table_addresses.RemoveRow(i_row_index, out o_error)) return false;

            return true;

        } // RemoveRow

        #endregion // Add and remove table rows


        #region Remove temporary files

        /// <summary>Remove all temporary used files
        /// <para>1. Delete backup files. Call of RemoveTemporaryUsedBackupFiles.</para>
        /// <remarks>Installer should create subdirectories Excel and Backups with modify rights</remarks>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool RemoveAllTemporaryUsedFiles(out string o_error)
        {
            o_error = "";

            if (!RemoveTemporaryUsedBackupFiles(out o_error))
            {
                return false;
            }

            if (!RemoveTemporaryUsedCheckinCheckoutFile(out o_error))
            {
                return false;
            }

            return true;
        }

        /// <summary>Remove temporary used backup files
        /// <para>1. Get backup file names. Call of Reset.GetBackupFiles.</para>
        /// <para>2. Delete backup files.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool RemoveTemporaryUsedBackupFiles(out string o_error)
        {
            o_error = "";

            string[] backup_file_names;
            if (!Reset.GetBackupFiles(out backup_file_names, out o_error))
            {
                return false;
            }
            try
            {
                foreach (string file_name in backup_file_names)
                {
                    File.Delete(file_name);
                }
            }

            catch (Exception e)
            {
                o_error = "Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }


            return true;

        } // RemoveTemporaryUsedBackupFiles

        /// <summary>Remove temporary used checkin-checkout file
        /// <para>1. Get checkin-checkout file name. Call of _GetCheckInOutLogFileName.</para>
        /// <para>2. Delete checkin-checkout file.</para>
        /// </summary>
        /// <param name="o_error">Error description</param>
        public bool RemoveTemporaryUsedCheckinCheckoutFile(out string o_error)
        {
            o_error = "";

            string checkin_checkout_local_file_name = _GetCheckInOutLogFileName();

            try
            {
                File.Delete(checkin_checkout_local_file_name);
            }

            catch (Exception e)
            {
                o_error = "Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            return true;

        } // RemoveTemporaryUsedCheckinCheckoutFile

        #endregion // Remove temporary files

    } // JazzMain

} // namespace
