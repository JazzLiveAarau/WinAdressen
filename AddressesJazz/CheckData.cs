using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using ExcelUtil;
using System.IO;
using Ftp;

namespace AddressesJazz
{
    /// <summary>Class with functions that check address data</summary>
    static public class CheckData
    {
        /// <summary>Check names.
        /// <para>Both first name and family name are not allowed to be empty.</para>
        /// </summary>
        /// <param name="i_first_name">First name (input string must be trimmed)</param>
        /// <param name="i_family_name">Family name (input string must be trimmed)</param>
        /// <param name="o_error">Error message</param>
        static public bool CheckNames(string i_first_name, string i_family_name, out string o_error)
        {
            o_error = "";
            bool data_is_ok = true;

            if (i_first_name == "" && i_family_name == "")
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoFirstNameNoFamilyName + "\n";
                data_is_ok = false;
            }

            if (data_is_ok)
            {
                return true;
            }
            else
            {
                return false;
            }
        } // CheckNames

        /// <summary>Check the Email address
        /// <para>There must be an @ in the E-Mail address.</para>
        /// </summary>
        /// <param name="i_email_address">Email address (input string must be trimmed)</param>
        /// <param name="o_error">Error message</param>
        static public bool CheckEmailAddress(string i_email_address, out string o_error)
        {
            o_error = "";
            bool data_is_ok = true;

            if (i_email_address != "")
            {
                if (!i_email_address.Contains("@"))
                {
                    o_error = AddressesJazzSettings.Default.ErrMsgEmailAddressNoAtSign;
                    data_is_ok = false;
                }
            }

            if (data_is_ok)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>Check that the Email address exists
        /// <para>(For the case that a Newsletter is requested).</para>
        /// </summary>
        /// <param name="i_email_address">Email address (input string must be trimmed)</param>
        /// <param name="o_error">Error message</param>
        static public bool EmailAddressExists(string i_email_address, out string o_error)
        {
            o_error = "";
            bool data_is_ok = true;

            if (i_email_address == "")
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoEmailAddressForNewsletter + "\n";
                data_is_ok = false;
            }

            if (data_is_ok)
            {
                return true;
            }
            else
            {
                return false;
            }
        } // EmailAddressExists

        /// <summary>Check the Mail address data
        /// <para>All address fields must be defined</para>
        /// </summary>
        /// <param name="i_street">Street (input string must be trimmed)</param>
        /// <param name="i_street_number">Street number (input string must be trimmed)</param>
        /// <param name="i_postal_code">Postal code (input string must be trimmed)</param>
        /// <param name="i_city">City (input string must be trimmed)</param>
        /// <param name="o_error">Error message</param>
        static public bool CheckMailAddress(string i_street, string i_street_number, string i_postal_code, string i_city, out string o_error)
        {
            o_error = "";
            bool data_is_ok = true;

            if (i_street == "")
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoStreetForMail + "\n";
                data_is_ok = false;
            }

            if (i_postal_code == "")
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoPostalCodeForMail + "\n";
                data_is_ok = false;
            }

            if (i_city == "")
            {
                o_error = AddressesJazzSettings.Default.ErrMsgNoCityForMail + "\n";
                data_is_ok = false;
            }


            if (data_is_ok)
            {
                return true;
            }
            else
            {
                return false;
            }
        } // CheckMailAddress

    } // class CheckData
}
