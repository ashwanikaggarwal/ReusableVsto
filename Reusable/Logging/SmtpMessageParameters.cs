using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Logging
{
    public class SmtpMessageParameters
    {
        public String Sender { get; set; }
        public String MachineName { get; set; }
        public String ToList { get; set; }
        public String CCList { get; set; }
        public String BccList { get; set; }
        public String Subject { get; set; }
        public String Body { get; set; }

        public override String ToString()
        {
            string convertedObject = string.Empty;
            convertedObject = String.Format("Sender: {0}, MachineName: {1}, Subject: {2}, ToList: {3}, CCList: {4}, BccList: {5}, Subject: {6}, Body: {7}", Sender, MachineName, Subject, ToList, CCList, BccList, Subject, Body);
            return convertedObject;
        }
    }
}
