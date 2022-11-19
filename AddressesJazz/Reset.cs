using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using ExcelUtil;
using System.IO;
using Ftp;
using AdressesUtility;

namespace AddressesJazz
{
    /// <summary>Class with functions that handle reset to an older (backup) address list
    /// <para></para>
    /// 
    /// </summary>
    static public class Reset
    {
        /// <summary>Download backup files from the server.
        /// <para>1. Host and user name from the configuration file (). Password () from JazzMain.m_ftp_password.</para>
        /// <para>2. The function checks if there is an Internet connection.</para>
        /// <para>3. The backup addresses CSV files are downloaded with function GetFiles in class Ftp.DownLoad.</para>
        /// </summary>
        /// <param name="o_error">Error message</param>
        static public bool DownloadBackupFiles(out string o_error)
        {
            o_error = "";

            string ftp_host = AddressesJazzSettings.Default.FtpHost;
            string ftp_user = AddressesJazzSettings.Default.FtpUser;
            string ftp_password = JazzMain.m_ftp_password;
            string exe_directory = JazzMain.m_exe_directory;

            if (!InternetUtil.IsInternetConnectionAvailable())
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoInternetConnection;

                return false;
            }

            Ftp.DownLoad ftp_download = new DownLoad(ftp_host, ftp_user, ftp_password);

            string server_address_directory = AddressesJazzSettings.Default.AddressesServerDir + AddressesJazzSettings.Default.AddressesBackupsDir + @"/";
            string local_address_directory = FileUtil.SubDirectory(AddressesJazzSettings.Default.AddressesBackupsDir, exe_directory);

            if (!ftp_download.GetFiles(server_address_directory, local_address_directory, out o_error))
            {
                return false;
            }

            return true;
        }


        /// <summary>Get available backup files as an array of strings.
        /// <para>Downloaded backup files are on a subdirectory (Backup) to the exe directory</para>
        /// <para>1. Function GetFilesDirectory in class FileUtil is called to get the file names.</para>
        /// </summary>
        /// <param name="o_backup_files">Array with backup file names</param>
        /// <param name="o_error">Error message</param>
        static public bool GetBackupFiles(out string[] o_backup_files, out string o_error)
        {
            ArrayList array_list_output = new ArrayList();
            o_backup_files = (string[])array_list_output.ToArray(typeof(string));
            o_error = "";

            string exe_directory = JazzMain.m_exe_directory;

            string[] file_extensions;
            ArrayList file_extensions_string_array = new ArrayList();
            file_extensions_string_array.Add(".csv");
            file_extensions = (string[])file_extensions_string_array.ToArray(typeof(string));

            string subdir_name = AddressesJazzSettings.Default.AddressesBackupsDir;
            string backup_directory = FileUtil.SubDirectory(subdir_name, exe_directory);
            bool reverse_array = true;


            if (!Directory.Exists(backup_directory))
            {
                o_error = backup_directory + " does not exist. Programming error";
                return false;
            }
            bool b_get = FileUtil.GetFilesDirectory(file_extensions, backup_directory, reverse_array, out o_backup_files);

            if (b_get)
            {
                return true;
            }
            else
            {
                o_error = AddressesJazzSettings.Default.ErrMsgGettingBackupFilesFailed;
                return false;
            }
        }


        /// <summary>Reset the addresses Table with a Table defined by a backup file.
        /// <para>1. Function JazzMain.ReplaceAddressesTable executes the command</para>
        /// </summary>
        /// <param name="i_main">JazzMain class</param>
        /// <param name="i_backup_file">Full name of the backup file</param>
        /// <param name="o_error">Error message</param>
        static public bool ResetWithBackupFile(JazzMain i_main, string i_backup_file, out string o_error)
        {
            o_error = "";

            if (!i_main.ReplaceAddressesTable(i_backup_file, out o_error))
            {
                return false;
            }


            return true;
        }

 

    } // class Reset
}
