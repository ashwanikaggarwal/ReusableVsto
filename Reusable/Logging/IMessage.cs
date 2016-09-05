using System;
namespace Reusable.Logging
{
    interface IMessage
    {
        string Description { get; set; }
        System.Diagnostics.EventLogEntryType EventType { get; set; }
        bool IsValid();
        string ToString();
        string User { get; set; }
        string Workstation { get; set; }
    }
}
