using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddressesJazz
{
    /// <summary>Internet utility functions</summary>
    public static class InternetUtil
    {
        #region Check internet connection

        /// <summary>Returns true if Internet connection is available.</summary>
        static public bool IsInternetConnectionAvailable()
        {
            string status_message = @"";

            string exe_directory = System.Windows.Forms.Application.StartupPath;

            JazzFtp.Input ftp_input = new JazzFtp.Input(exe_directory, JazzFtp.Input.Case.CheckInternetConnection);

            JazzFtp.Result ftp_result = JazzFtp.Execute.Run(ftp_input);

            if (!ftp_result.Status)
            {
                // Programming error
                status_message = @"IsInternetConnectionAvailable JazzFtp.Execute.Run failed " + ftp_result.ErrorMsg;

                MessageBox.Show(status_message);

                return false;
            }

            return ftp_result.BoolResult;

        } // IsInternetConnectionAvailable

        #endregion // Check internet connection
    } // InternetUtil

} // namespace
