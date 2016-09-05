using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.ComponentModel;using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Logging
{
    public class SmtpMailMessage
    {
        private SmtpMessageParameters parameters;
        public SmtpMessageParameters Parameters
        {
            get { return parameters; }
            set { parameters = value; }
        }

        public SmtpMailMessage(SmtpMessageParameters settings)
        {
            parameters = settings;
        }

        static bool mailSent = false;
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
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

        public void Send()
        {
            SmtpClient client = new SmtpClient("SMTP.wanlink.us", 25);

            MailAddress from = new MailAddress(parameters.Sender);
            MailAddress to = new MailAddress(parameters.ToList);
            
            MailMessage message = new MailMessage(from, to);
            message.Body = parameters.Body;
            message.BodyEncoding =  System.Text.Encoding.UTF8;
            
            message.Subject = parameters.Subject;
            message.SubjectEncoding = System.Text.Encoding.UTF8;

            // Set the method that is called back when the send operation ends.
            client.SendCompleted += new 
            SendCompletedEventHandler(SendCompletedCallback);
            
            // The userState can be any object that allows your callback  
            // method to identify this send operation. 
            string userState = parameters.ToString();
            client.SendAsync(message, userState);

            // Clean up.
            //message.Dispose();
        }
    }
}
