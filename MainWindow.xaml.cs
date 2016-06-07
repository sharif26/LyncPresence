/*=====================================================================
  This file is part of the Microsoft Unified Communications Code Samples.

  Copyright (C) 2012 Microsoft Corporation.  All rights reserved.

This source code is intended only as a supplement to Microsoft
Development Tools and/or on-line documentation.  See these other
materials for detailed information regarding Microsoft code samples.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
=====================================================================*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Threading.Tasks;
using Microsoft.Lync.Model;
using System.Net.Mail;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Timers;

namespace PresencePublication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
 

    public partial class MainWindow : Window
    {
        #region Fields
        // Current dispatcher reference for changes in the user interface.
        private Dispatcher dispatcher;
        private LyncClient lyncClient;
        private Contact peer, peer2;
        private bool isAwayM = false;
        private bool isAwayS = false;
        private bool isAwayBoth = false;
        private DateTime dtm = DateTime.Now;
        private DateTime dts = DateTime.Now;
//        private Timer aTimer;
        

        #endregion

        public MainWindow()
        {
            InitializeComponent();

            //Save the current dispatcher to use it for changes in the user interface.
            dispatcher = Dispatcher.CurrentDispatcher;
        }

        #region Handlers for user interface controls events
        /// <summary>
        /// Handler for the Loaded event of the Window.
        /// Used to initialize the values shown in the user interface (e.g. availability values), get the Lync client instance
        /// and start listening for events of changes in the client state.
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Add the availability values to the ComboBox
            availabilityComboBox.Items.Add(ContactAvailability.Free);
            availabilityComboBox.Items.Add(ContactAvailability.Busy);
            availabilityComboBox.Items.Add(ContactAvailability.DoNotDisturb);
            availabilityComboBox.Items.Add(ContactAvailability.Away);

            //Listen for events of changes in the state of the client
            try
            {
                lyncClient = LyncClient.GetClient();
            }
            catch (ClientNotFoundException clientNotFoundException)
            {
                Console.WriteLine(clientNotFoundException);
                return;
            }
            catch (NotStartedByUserException notStartedByUserException)
            {
                Console.Out.WriteLine(notStartedByUserException);
                return;
            }
            catch (LyncClientException lyncClientException)
            {
                Console.Out.WriteLine(lyncClientException);
                return;
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                    return;
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            lyncClient.StateChanged +=
                new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

//            SetAvailability();
            //Update the user interface
            UpdateUserInterface(lyncClient.State);
//            SendEmail("test");
//            SendSMS("Hello");

//            SetTimer();
//            System.Threading.Thread.Sleep(5000);


        }
/*
        private void SetTimer()
        {
            // Create a timer with a two second interval.
            aTimer = new Timer(60000);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;
            aTimer.Start();
            //aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }
*/
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            Console.WriteLine("The Elapsed event was raised at {0:HH:mm:ss.fff}", e.SignalTime);

            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.ActivityId, "Available");      //ContactAvailability.Free       availabilityComboBox.SelectedItem

            try
            {
                //ContactAvailability currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                string activity = (string) lyncClient.Self.Contact.GetContactInformation(ContactInformationType.ActivityId);
                Console.WriteLine("Before: " + activity);

                IAsyncResult result = lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }


        }

        /// <summary>
        /// Handler for the SelectionChanged event of the Availability ComboBox. Used to publish the selected availability value in Lync
        /// </summary>
        private void AvailabilityComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Add the availability to the contact information items to be published
            Dictionary<PublishableContactInformationType, object> newInformation =
                new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.Availability, availabilityComboBox.SelectedItem);

            //Publish the new availability value
            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation,PublishContactInformationCallback, null);
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        /// <summary>
        /// Handler for the Click event of the Note Button. Used to publish a new personal note value in Lync
        /// </summary>
        private void SetNoteButton_Click(object sender, RoutedEventArgs e)
        {
//            SendKeys.Send("{CAPSLOCK}");


            //Add the personal note to the contact information items to be published
            Dictionary<PublishableContactInformationType, object> newInformation =
                new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.PersonalNote, personalNoteTextBox.Text);

            //Publish the new personal note value
            try
            {
                lyncClient.Self.BeginPublishContactInformation(newInformation,PublishContactInformationCallback, null);
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
           


        }

        /// <summary>
        /// Handler for the Click event of the SignInOut Button. Used to sign in or out Lync depending on the current client state.
        /// </summary>
        private void SignInOutButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (lyncClient.State == ClientState.SignedIn)
                {
                    //Sign out If the current client state is Signed In
                    lyncClient.BeginSignOut(SignOutCallback, null);
                }
                else if (lyncClient.State == ClientState.SignedOut)
                {
                    //Sign in If the current client state is Signed Out
                    lyncClient.BeginSignIn(null, null, null, SignInCallback, null);
                }
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }
        #endregion

        #region Handlers for Lync events
        /// <summary>
        /// Handler for the ContactInformationChanged event of the contact. Used to update the contact's information in the user interface.
        /// </summary>
        private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            //Only update the contact information in the user interface if the client is signed in.
            //Ignore other states including transitions (e.g. signing in or out).
            if (lyncClient.State == ClientState.SignedIn)
            {
                //Get from Lync only the contact information that changed.

                if (e.ChangedContactInformation.Contains(ContactInformationType.DisplayName))
                {
                    //Use the current dispatcher to update the contact's name in the user interface.
                    dispatcher.BeginInvoke(new Action(SetName));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    //Use the current dispatcher to update the contact's availability in the user interface.
                    dispatcher.BeginInvoke(new Action(SetAvailability));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.PersonalNote))
                {
                    //Use the current dispatcher to update the contact's personal note in the user interface.
                    dispatcher.BeginInvoke(new Action(SetPersonalNote));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Photo))
                {
                    //Use the current dispatcher to update the contact's photo in the user interface.
                    dispatcher.BeginInvoke(new Action(SetContactPhoto));
                }
            }
        }

        private void PeerContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            //Only update the contact information in the user interface if the client is signed in.
            //Ignore other states including transitions (e.g. signing in or out).
            /*
            if (lyncClient.State == ClientState.SignedIn)
            {
                //Get from Lync only the contact information that changed.

                if (e.ChangedContactInformation.Contains(ContactInformationType.DisplayName))
                {
                    //Use the current dispatcher to update the contact's name in the user interface.
                    dispatcher.BeginInvoke(new Action(SetName));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    //Use the current dispatcher to update the contact's availability in the user interface.
                    dispatcher.BeginInvoke(new Action(SetAvailability));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.PersonalNote))
                {
                    //Use the current dispatcher to update the contact's personal note in the user interface.
                    dispatcher.BeginInvoke(new Action(SetPersonalNote));
                }
                if (e.ChangedContactInformation.Contains(ContactInformationType.Photo))
                {
                    //Use the current dispatcher to update the contact's photo in the user interface.
                    dispatcher.BeginInvoke(new Action(SetContactPhoto));
                }
            }
            */
            if (lyncClient.State == ClientState.SignedIn)
            {
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    ContactAvailability pca = (ContactAvailability)peer.GetContactInformation(ContactInformationType.Availability);
                    //Console.WriteLine(pca);
                    if (pca == ContactAvailability.Away || pca == ContactAvailability.Offline) 
                    { 
                        //SendSMS("Away");
                        isAwayS = true;
                    }
                    //if (pca == ContactAvailability.Free) 
                    else
                    { 
                        //SendSMS("Online");
                        isAwayS = false;

                        if (isAwayBoth)
                        {
                            SendSMS("S online");
//                            SendEmail("S online");
                            isAwayBoth = false;
                        }
                    }
                    if (!dts.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString()))
                    {
                        Console.WriteLine("S: " + pca + ",AwayS? " + isAwayS + ", Time=" + DateTime.Now.ToString("h:mm:ss.fffffff tt"));
                        if (isAwayM && isAwayS)
                        {
                            Console.WriteLine("Both are away");
                            isAwayBoth = true;
                            SendSMS("Both are away");
//                            SendEmail("Both are away");
                        }
                    }
                    dts = DateTime.Now;
                    
                    //SendEmail("" + pca.ToString());
                }
            }
        }

        private void Peer2Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            if (lyncClient.State == ClientState.SignedIn) 
            {
                if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
                {
                    ContactAvailability pca = (ContactAvailability)peer2.GetContactInformation(ContactInformationType.Availability);
                    if (pca == ContactAvailability.Away) 
                    { 
                        //SendSMS("Away"); 
                        isAwayM = true;
                    }
                    //if (pca == ContactAvailability.Free) 
                    else
                    { 
                        //SendSMS("Online"); 
                        isAwayM = false;

                        if (isAwayBoth)
                        {
                            SendSMS("M online");
//                            SendEmail("M online");
                            isAwayBoth = false;
                        }
                    }
                    if (!dtm.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString()))
                    {
                        Console.WriteLine("M: " + pca + ",AwayM? " + isAwayM + ", Time=" + DateTime.Now.ToString("h:mm:ss.fffffff tt"));
                        if (isAwayM && isAwayS)
                        {
                            Console.WriteLine("Both are away");
                            isAwayBoth = true;
                            SendSMS("Both are away");
//                            SendEmail("Both are away");
                        }
                    }
                    dtm = DateTime.Now;
                    
                    //SendEmail("" + pca.ToString());
                }
            }
        }

/*        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }
*/
        private void SendEmail( string msg ) 
        {
            // Command line argument must the the SMTP host.        sahmed@chesterfield.mo.us
            MailAddress to = new MailAddress("sahmed@chesterfield.mo.us");
            //MailAddress to = new MailAddress("shiarif@gmail.com");
            MailAddress from = new MailAddress("pvtd@chesterfield.mo.us");
            //MailAddress from = new MailAddress("shiarif@gmail.com");
            MailMessage message = new MailMessage(from, to);
            message.Subject = "User is: " + msg;
            message.Body = @"Using this new feature, you can send an e-mail message from an application very easily.";
            // Use the application or machine configuration to get the 
            // host, port, and credentials.
            SmtpClient client = new SmtpClient("email.chesterfield.mo.us");
            //SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
            //client.Credentials = new NetworkCredential("shiarif@gmail.com", "");
            //client.EnableSsl = true;
            Console.WriteLine("Sending an e-mail message to {0} at {1} by using the SMTP host={2}.",
                to.User, to.Host, client.Host);
            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in CreateTestMessage3(): {0}", ex.ToString());
            }
 
        }

        private void SendSMS( string text ) 
        {

            try
            {
                // Create a request using a URL that can receive a post. 
                WebRequest request = WebRequest.Create("https://www.google.com/voice/b/0/sms/send/");
                // Set the Method property of the request to POST.
                request.Method = "POST";
                request.Proxy = new WebProxy("http://192.168.10.2:8080", true, null, CredentialCache.DefaultCredentials);

                //request.Headers.Add("", "");
                request.Headers.Add("origin", "https://www.google.com");
                request.Headers.Add("accept-encoding", "gzip, deflate");
                request.Headers.Add("accept-language", "en-US,en;q=0.8,bn;q=0.6");
//                request.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.109 Safari/537.36");
//                request.Headers.Add("accept", "*/*");
//                request.Headers.Add("referer", "https://www.google.com/voice/b/0");
                request.Headers.Add("authority", "www.google.com");
//                request.Headers.Add("cookie", "gv=DQAAABkBAACsEqx1OpupDSjkcY3VPnVrErJMiB2K0kB3otoFtm5hO3d3XrFYGHnqLksoWOrK6asFmn9goZju1iIhyi0MoAaHQ8_6TuC9UQoaPMybA6gU4I2exCDw_Huep994aP-22ulBQcyCPZQqAisjze1Stb_pqaQj2bR6bHJ_U3MtbnD1kPhVwO2VKCEucGOSdi5ATBdS6xuTuAz0UfLdmSmqJMD487vtV-WAH1qKXs47Bqw9sgsiAbGzeZYKt1_GncsDXj6aOhy8ggM5wxaRsM7cXBWReN9x7ZeaEpCqoJ4bJ2x9vEk-Hmzw_I_duHDmk-g3eulYUJ5wM8o-Vq7HoOUWAqC22XQ4k0YpEw4sfLlRyhHhA5mnewBjtaMpKx_a-6LPD7A; GMAIL_RTT=30; SID=DQAAABcBAACWB23GkBaekFv__Puscz1qOFo6scjaGVgmzaIgMBMnIlclD0hiv-uq2kktrNkDRCq5QStNQBxcK58w93jWhOWYzfJdCsPuf2Sjwy79ARpEaRcIqduCiZLTcM5M8DDdC5o2XEhMk12TDKn3Kb0E1NtAEC6Jm-sF_wlLI5c6LdxinQtydLp3mDB0P_i6_VUo22jcFCRtsq3RKOQZSmxqtvSIXjYNBCw-SgzCOmgGkaso2tGSHTYk6PksA_hC322A63t9-zZQdXcLw27GoTsaHBwfziWyFNNJI8PR_frrmsVc8B5fLa6Qj2R6MKWUwUTnZ-CKCphP0_5eWzlRRq_TmPvODAX1gnVgUul7Tg8pK3e1Jd3Xpxk_LierEWF17zYJAfE; HSID=AFtyR9N0TDkzIiUNJ; SSID=AIs-owof1hcFZC_fG; APISID=hhQASfDc_4wJsPuY/AVsbvzjZLbTmZPikE; SAPISID=EV1WT_8J0k2ttLCO/AL5NYIqqGeyqVi73s; NID=78=SbBi4DGNmgnwyCLoL8SP0qirzc10NQSCAMe8yTgIZuPuCDMlzG3d3kEgQ68jmoSTzP08cyUTIThsDDygzfCphlk9b6Lc_o9HFIubi25hMfKgDFZYYv_2pST37z1YRRa6jOA5DUj9JREoEdSfDRkkIzMuw3phEnSPsbwlDwc1JE8q_qpv7nlzp_LlY1iNz7lx73yM6vvK0onX; S=grandcentral=vxGYgHfWuYe9JH7knwfrmw; _ga=GA1.1.331231560.1460046787; _gat=1");
                request.Headers.Add("cookie", "gv=RAPedRgPlHnmTkiHqg4Zy63ebNhRtO4KxHqPWM0T64h4-_cNGcadUidriFt0mp6q-jslkA.; GMAIL_RTT=28; SID=RAPedX6vm8ZsR8lDeqGBzfGlEfTP3NJvvaC5gh5pyVPR1UbA-d4cMy9_ZhvufHgmO6FpWQ.; HSID=A1UF4MF2g6YxGTLMZ; SSID=A_9YmujmhwMKGuy2y; APISID=CHDdJgDgmkNHE7Wq/Akjn8KhUoE-MuHK28; SAPISID=oGVtA_6tSoZlFVA1/A1VLjslvZMvp93EmJ; NID=79=aBquJKxgpBePiG0AQ9JuHm8nnxxzKxozvgQIrPGKJB3uATy9evDv9mb3eti0y8Bu1awVSHW8K0NS1mChDtVGsVDVmsuC2E2NtAGa8GvW7J4nGklA0aGYDRKQQaE6PWPPi8edYCq4bkfjMWRfDzav0gubVLEm8RMLwCrYU7ynqMCNNoo0-N-5MwgaKl2a7zBWYRd9XKiCyUu9; _ga=GA1.1.1424887623.1462985744; _gat=1; S=grandcentral=zj3HARSR5yOK4k36lvonUw");
//                request.Headers.Add("x-client-data", "CKW2yQEIxLbJAQi1lMoBCP2VygE=");

                // Create POST data and convert it to a byte array.
                string postData = "id=9bb534e53c1ca5b97d64290d5c6f2abbf1cf2856&phoneNumber=%2B15734790167&conversationId=9bb534e53c1ca5b97d64290d5c6f2abbf1cf2856&text=" +
                    text+"&contact=Sharif%20Ahmed&_rnr_se=ld0M4R%2FX5GjhRJxxEJIij1lb3AI%3D";
                byte[] byteArray = Encoding.UTF8.GetBytes(postData);
                // Set the ContentType property of the WebRequest.
                request.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
                // Set the ContentLength property of the WebRequest.
                request.ContentLength = byteArray.Length;
                // Get the request stream.
                Stream dataStream = request.GetRequestStream();
                // Write the data to the request stream.
                dataStream.Write(byteArray, 0, byteArray.Length);
                // Close the Stream object.
                dataStream.Close();
                // Get the response.
                WebResponse response = request.GetResponse();
                // Display the status.
                Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                // Get the stream containing content returned by the server.
                dataStream = response.GetResponseStream();
                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);
                // Read the content.
                string responseFromServer = reader.ReadToEnd();
                // Display the content.
//                Console.WriteLine(responseFromServer);
                // Clean up the streams.
                reader.Close();
                dataStream.Close();
                response.Close();
            }
            catch (Exception e) {
                Console.WriteLine(e.ToString());
            }

        }

        /// <summary>
        /// Handler for the StateChanged event of the contact. Used to update the user interface with the new client state.
        /// </summary>
        private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            //Use the current dispatcher to update the user interface with the new client state.
            dispatcher.BeginInvoke(new Action<ClientState>(UpdateUserInterface), e.NewState);
        }
        #endregion

        #region Callbacks
        /// <summary>
        /// Callback invoked when LyncClient.BeginSignIn is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void SignInCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignIn(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        /// <summary>
        /// Callback invoked when LyncClient.BeginSignOut is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void SignOutCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignOut(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        /// <summary>
        /// Callback invoked when Self.BeginPublishContactInformation is completed
        /// </summary>
        /// <param name="result">The status of the asynchronous operation</param>
        private void PublishContactInformationCallback(IAsyncResult result)
        {
            lyncClient.Self.EndPublishContactInformation(result);
            
            //ContactAvailability currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
            string activity = (string)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.ActivityId);
            Console.WriteLine("After: " + activity + ", Result: " + result.IsCompleted);
        }
        #endregion

        /// <summary>
        /// Updates the user interface
        /// </summary>
        /// <param name="currentState"></param>
        private void UpdateUserInterface(ClientState currentState)
        {
            //Update the client state in the user interface
            clientStateTextBox.Text = currentState.ToString();

            if (currentState == ClientState.SignedIn)
            {
                //Listen for events of changes of the contact's information
                lyncClient.Self.Contact.ContactInformationChanged +=
                    new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);

                peer = lyncClient.ContactManager.GetContactByUri("sip:sdecker@chesterfield.mo.us");
                peer.ContactInformationChanged += new EventHandler<ContactInformationChangedEventArgs>(PeerContact_ContactInformationChanged);
                ContactAvailability pca = (ContactAvailability)peer.GetContactInformation(ContactInformationType.Availability);
                if (pca == ContactAvailability.Away ||pca == ContactAvailability.Offline) { isAwayS = true; }
                Console.WriteLine("S: " + pca + ",isAwayS: " + isAwayS);
                peer2 = lyncClient.ContactManager.GetContactByUri("sip:mhaug@chesterfield.mo.us");
                peer2.ContactInformationChanged += new EventHandler<ContactInformationChangedEventArgs>(Peer2Contact_ContactInformationChanged);
                pca = (ContactAvailability)peer2.GetContactInformation(ContactInformationType.Availability);
                if (pca == ContactAvailability.Away) { isAwayM = true; }
                Console.WriteLine("M: " + pca + ",isAwayM: " + isAwayM);
                if (isAwayM && isAwayS)
                {
                    Console.WriteLine("Both are away");
                    isAwayBoth = true;
                    //SendSMS("Both are away");
                }

                //Get the contact's information from Lync and update with it the corresponding elements of the user interface.
                SetName();
                SetAvailability();
                SetPersonalNote();
                SetContactPhoto();

                //Update the SignInOut button content
                signInOutButton.Content = "Sign Out";

                //Enable elements in the user interface
                personalNoteTextBox.IsEnabled = true;
                availabilityComboBox.IsEnabled = true;
                setNoteButton.IsEnabled = true;
            }
            else
            {
                //Update the SignInOut button content
                signInOutButton.Content = "Sign In";

                //Disable elements in the user interface
                personalNoteTextBox.IsEnabled = false;
                availabilityComboBox.IsEnabled = false;
                setNoteButton.IsEnabled = false;

                //Change the color of the border containing the contact's photo to match the contact's offline status
                availabilityBorder.Background = Brushes.LightSlateGray;
            }
        }

        /// <summary>
        /// Gets the contact's current availability value from Lync and updates the corresponding elements in the user interface
        /// </summary>
        private void SetAvailability()
        {
            //Get the current availability value from Lync
            ContactAvailability currentAvailability = 0;
            currentAvailability = ContactAvailability.Free;
            
            try
            {
                currentAvailability = (ContactAvailability)
                                                          lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            if (currentAvailability != 0)
            {
                //Update the availability ComboBox with the contact's current availability.
                availabilityComboBox.SelectedValue = currentAvailability;

                //Choose a color to match the contact's current availability and update the border area containing the contact's photo
                Brush availabilityColor;
                switch (currentAvailability)
                {
                    case ContactAvailability.Away:
                        availabilityColor = Brushes.Yellow;
                        break;
                    case ContactAvailability.Busy:
                        availabilityColor = Brushes.Red;
                        break;
                    case ContactAvailability.BusyIdle:
                        availabilityColor = Brushes.Red;
                        break;
                    case ContactAvailability.DoNotDisturb:
                        availabilityColor = Brushes.DarkRed;
                        break;
                    case ContactAvailability.Free:
                        availabilityColor = Brushes.LimeGreen;
                        break;
                    case ContactAvailability.FreeIdle:
                        availabilityColor = Brushes.LimeGreen;
                        break;
                    case ContactAvailability.Offline:
                        availabilityColor = Brushes.LightSlateGray;
                        break;
                    case ContactAvailability.TemporarilyAway:
                        availabilityColor = Brushes.Yellow;
                        break;
                    default:
                        availabilityColor = Brushes.LightSlateGray;
                        break;
                }
                availabilityBorder.Background = availabilityColor;
            }

/*
            Dictionary<PublishableContactInformationType, object> newInformation = new Dictionary<PublishableContactInformationType, object>();
            newInformation.Add(PublishableContactInformationType.Availability, ContactAvailability.Free);      //ContactAvailability.Free       availabilityComboBox.SelectedItem

            try
            {
                currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                Console.WriteLine("Before: "+currentAvailability);

                IAsyncResult result = lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);


            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
*/            
        }

        /// <summary>
        /// Gets the contact's name from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetName()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.DisplayName)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            nameTextBlock.Text = text;
        }

        /// <summary>
        /// Gets the contact's personal note from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetPersonalNote()
        {
            string text = string.Empty;
            try
            {
                text = lyncClient.Self.Contact.GetContactInformation(ContactInformationType.PersonalNote)
                              as string;
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            personalNoteTextBox.Text = text;
        }

        /// <summary>
        /// Gets the contact's photo from Lync and updates the corresponding element in the user interface
        /// </summary>
        private void SetContactPhoto()
        {
            try
            {
                using (Stream photoStream =
                    lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Photo) as Stream)
                {
                    if (photoStream != null)
                    {
                        BitmapImage bm = new BitmapImage();
                        bm.BeginInit();
                        bm.StreamSource = photoStream;
                        bm.EndInit();
                        photoImage.Source = bm;
                    }
                }
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }

        /// <summary>
        /// Identify if a particular SystemException is one of the exceptions which may be thrown
        /// by the Lync Model API.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }
    }
}
