using System;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using EASendMail;

//=====================[Read This]===========================
//    if you are reading this, it means that you have got into my source code. Not hard to do.
//
//    But this is my code, and mine alone.
//    It has been developed by me, is my intellectual property and is what I use to feed my family.
//    So what you are doing is stealing from me, and worse, you are stealing from my family.

//    So here's the thing.
//
//    If I catch you, and believe me this wont be hard, I will fucking crucify you, the company you work for, and if you are a contractor, I will take everything I can from you.
//
//    Take this seriously.
//
//    Trimble is a $2bn US company with 6000 employees. They tried to screw me.
//    I took them to court.
//    Me, on my own.
//    I won.
//    6 Figure $ settlement thank you.
//    You willing to take that chance? You got that sort of money to defend your stealing code?
//
//    People think I am this pleasant smiling guy from South Africa, always willing to help.
//    Yeah yeah whatever, until someone crosses me.
//    Then my retribution knows no bounds, and I will aggressively and actively persue and destroy anyone that touches me or my family.

//    Dont steal from me. Dont even contemplate such a thing.
//    Close this code and delete it from your computer, your memory stick, your mind, your scraps of notes, wherever you have uploaded it, and go have a coffee.
//    Forget you even saw this.

//    The alternative is not worth it.
//
//     Trust me on that one....... as I said....I will fucking crucify you.
//     Stone fucking dead.
//
//=======================================================================================================================================
//
//
// To add this to a project..
// Right Click References
//  add the reference
//  add Reference.. browse to LicenseLibrary from the bin/Release location or use the Recent option
// include after static void Main(): 
//      gnaLicenseLibrary gna = new gnaLicenseLibrary();
//
// The versions of the packages must match or you get an error message that the version could not be found.
// Update so that the packages all match
//
// Remember to recompile LicenseLibrary using Release after any change
//






namespace gnaLicenseLibrary
{
    public class LicenseLibrary
    {

#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416




        public void WelcomeMessage(string strMessage)
        {
            Console.Title = strMessage;

            Console.WriteLine(" ");
            Console.WriteLine("GNA Geomatics software");
            Console.WriteLine("Julian Gray");
            Console.WriteLine("+49 176 7299 7904");
            Console.WriteLine("gna.geomatics@gmail.com");
            Console.WriteLine(" ");
            Console.WriteLine("Do not steal my software.");
            Console.WriteLine("-------------------------");
            Console.WriteLine(" ");
            Console.WriteLine("");
        }


        public string checkLicenseStatus(string strProject, string strEmailLogin, string strEmailPassword, string strSendEmail)
        {


            string strStatus = "empty";

            //
            // Check the number of days remaining on the license
            // The start date and validity period must be set using the software "SetSoftwareKey"
            //

            // This module must be at the start of all software modules
            // use:
            //
            // string strSendEmail = "No";
            // string strSoftwareKey = gna.checkSoftwareReferenceDate(strProject, strEmailLogin, strEmailPassword, strSendEmail);
            // if (strSoftwareKey == "expired") goto TheEnd;
            //
            // TheEnd:
            //  Console.WriteLine("Done");
            //
            // the answer is either "valid" or "expired"
            //  strSendEmail = "Yes" when CheckSoftwareKey is set as a scheduled task 
            //  strSendEmail is set to "No" when used inside the software
            //

            //Console.WriteLine("checkSoftwareReferenceDate");
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Diebus");
            string strValidityPeriod = key.GetValue("TempusValide", "No Value").ToString();
            key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Portunus");
            string strReferenceDate = key.GetValue("Clavis", "No Value").ToString();

            Console.WriteLine("Licence period: " + strValidityPeriod + " days");
            Console.WriteLine("Reference date: " + strReferenceDate);

            DateTime InstallDate = DateTime.ParseExact(strReferenceDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            DateTime TodayDate = DateTime.Today;

            TimeSpan interval = (TodayDate - InstallDate);
            int iRemainingDays = Convert.ToInt16(strValidityPeriod) - interval.Days;

            Console.WriteLine("Remaining days: " + iRemainingDays.ToString() + " days");

            if (iRemainingDays < 0)
            {
                strStatus = "expired";

                Console.WriteLine(" ");
                Console.WriteLine("The software license has expired.");
                Console.WriteLine("Please contact Julian Gray on +49 176 7299 7904 to reactivate");
                Console.WriteLine(" ");
                Console.ReadLine();
            }
            else
            {
                strStatus = "valid";
            }

            if ((iRemainingDays < 4) & (strSendEmail == "Yes"))
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3");
                    {
                        oMail.From = "gna.geomatics@gmail.com";
                        oMail.To = new AddressCollection("gna.geomatics@gmail.com");
                    };

                    // Set email subject
                    oMail.Subject = "Software license about to expire (" + strProject + ")";
                    // Set email body
                    oMail.TextBody = "No of days remaining: " + iRemainingDays.ToString();

                    // SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.gmail.com")
                    {
                        User = strEmailLogin,
                        Password = strEmailPassword,
                        ConnectType = SmtpConnectType.ConnectTryTLS,
                        Port = 587
                    };

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                    Console.WriteLine("Advisory email issued");

                }
                catch (Exception ep)
                {
                    Console.WriteLine("Email transmission failed..");
                    Console.WriteLine(strEmailLogin);
                    Console.WriteLine(strEmailPassword);
                    Console.WriteLine("");
                    Console.WriteLine(ep.Message);
                    Console.ReadKey();
                }
            }

            return strStatus;
        }



    }
}
