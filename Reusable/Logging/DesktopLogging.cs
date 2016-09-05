using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace Reusable.Logging
{
    # region Log-specific Properties and Enums

    /// <summary>
    /// Enumeration for error levels, correspond to Enterprise logging values
    /// </summary>
    [FlagsAttribute]
    public enum LogLevelParameters
    {
        //0
        None = 0,

        //System.Diagnostics.TraceEventType.Critical
        Critical = 1,

        //System.Diagnostics.TraceEventType.Error
        Error = 2,

        //System.Diagnostics.TraceEventType.Warning
        Warning = 4,

        //System.Diagnostics.TraceEventType.Information
        Information = 8,

        //System.Diagnostics.TraceEventType.Verbose
        Verbose = 16,

        //System.Diagnostics.TraceEventType.Start
        Start = 256,

        //System.Diagnostics.TraceEventType.Stop
        Stop = 512,

        //System.Diagnostics.TraceEventType.Transfer
        Logical = 1024
    }

    # endregion

    public class DesktopLogging : IDisposable
    {
        /// <summary>
        /// Enumeration defines the type of log the class instantiates
        /// </summary>
        public enum LogType
        {
            EventViewer,
            TextFile
        }

        # region Constructors

        /// <summary>
        /// constructor for logging class
        /// </summary>
        /// <param name="file"></param>
        /// <param name="logtype"></param>
        public DesktopLogging(string file, LogType type)
        {
            logWriterStart(file, type);
        }

        /// <summary>
        /// callable method that constructs class
        /// </summary>
        /// <param name="file"></param>
        /// <param name="logtype"></param>
        private void logWriterStart(string file, LogType type)
        {
            _LogType = type;

            if (_LogType == LogType.TextFile)
            {
                try
                {
                    _Stream = new FileStream(file, FileMode.Create, FileAccess.Write);
                    _OutputPath = file;
                    _Result = new StreamWriter(_Stream);
                    _Result.Close();
                }
                catch
                {
                    //empty catch in case of conflicts
                    //potential issues
                    // - no directory rights
                    // - no space
                    // - invalid directory
                }
                finally
                {
                    _OutputPath = file;
                }
            }
            else
            {
                _OutputPath = file;
            }

            //sets default as error, to avoid user errors
            _LogLevel = LogLevelParameters.Error;
        }

        # endregion

        # region Properties

        //general class variables
        private LogLevelParameters _LogLevel;
        private LogType _LogType;

        //text logging componenets
        private StreamWriter _Result;
        private FileStream _Stream;

        private string _OutputPath;
        /// <summary>
        /// The text file set as default target for text logging
        /// </summary>
        public string OutputPath
        {
            get { return _OutputPath; }
            set { _OutputPath = value; }
        }

        private string _EventSource;
        /// <summary>
        /// Defines the event source for errors 
        /// </summary>
        public string EventSource
        {
            get { return _EventSource; }
            set { _EventSource = value; }
        }

        private string _EventLogName;
        /// <summary>
        /// Defines the event log to write to in EventViewer
        /// </summary>
        public string EventLogName
        {
            get { return _EventLogName; }
            set { _EventLogName = value; }
        }

        /// <summary>
        /// Log level when logging class is created/set provides threshold for logging.
        /// If value is one of the valid critieris use it
        /// else if lower than 1, set to 2,
        /// else if greater than 4096, set to 4096
        /// else assigns to closet number
        /// </summary>
        public LogLevelParameters LogLevel
        {
            get { return _LogLevel; }
            set
            {
                Int16 level = Convert.ToInt16(value, CultureInfo.InvariantCulture);

                if (level == 1 || level == 2 || level == 4 || level == 8 || level == 16 || level == 1024)
                {
                    _LogLevel = value;
                }
                else if (value < LogLevelParameters.Critical)
                {
                    _LogLevel = LogLevelParameters.Error;
                }
                else if (value > LogLevelParameters.Logical)
                {
                    _LogLevel = LogLevelParameters.Logical;
                }
                else
                {
                    _LogLevel = LogLevelParameters.Logical;
                }
            }
        }

        # endregion

        # region Methods

        /// <summary>
        /// Formatter for text message, prepends time stamp to each message
        /// </summary>
        /// <param name="message"></param>
        /// <param name="level"></param>
        /// <returns></returns>
        private static string FormatMessage(string message, string level)
        {
            string messageFormatted = string.Format(CultureInfo.InvariantCulture, "{0} || {1} || {2}", System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), level, message);
            return messageFormatted;
        }

        /// <summary>
        /// Primary method for writing messages to text
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="level"></param>
        public void WriteLog(string msg, LogLevelParameters level)
        {
            lock (this)
            {
                if (level <= _LogLevel || level == LogLevelParameters.Start || level == LogLevelParameters.Stop)
                {
                    if (_LogType == LogType.TextFile)
                    {
                        WriteLogToText(msg, level);
                        if (level == LogLevelParameters.Error)
                        {
                            WriteLogToEventViewer(_EventSource, _EventLogName, msg);
                        }
                    }
                    else
                    {
                        WriteLogToEventViewer(_EventSource, _EventLogName, msg);
                    }
                }
            }
        }

        /// <summary>
        /// Primary private method for writing messages to the event viewer
        /// </summary>
        /// <param name="NewLine"></param>
        /// <param name="level"></param>
        private void WriteLogToText(string newLine, LogLevelParameters level)
        {
            try
            {
                string message = FormatMessage(newLine, level.ToString());

                _Stream = new FileStream(_OutputPath, FileMode.Append, FileAccess.Write);
                _Result = new StreamWriter(_Stream);
                _Result.WriteLine(message);
                _Result.Close();
                _Stream = null;
            }
            catch
            {
                //empty catch, in case of file conflict or right sissue
            }
        }

        /// <summary>
        /// Primary method for writing event viewer messages. 
        /// Marked as static and does not require instantiation
        /// </summary>
        /// <param name="NewLine"></param>
        /// <param name="level"></param>
        public static void WriteLogToEventViewer(string sourceName, string logName, string newLine)
        {
            string sEvent = string.Empty;
            int logEventId = 4000;

            try
            {
                sEvent = string.Format(CultureInfo.InvariantCulture, "{0}", newLine);

                if (!EventLog.SourceExists(sourceName))
                {
                    EventLog.CreateEventSource(sourceName, logName);
                }

                EventLog.WriteEntry(sourceName, sEvent, EventLogEntryType.Error, logEventId);
            }
            catch (ArgumentException ex)
            {
                EventLog.WriteEntry(ex.Message, sEvent, EventLogEntryType.Error, 0);
            }
            catch
            {
                //empty catch to avoid problems with Citrix
                //Citrix clients are unable to write to event viewer

                //also empty in case there are unexpected issues writing to event viewer
            }
        }

        # endregion

        # region Dispose

        /// <summary>
        /// Dispose() calls Dispose(true) and GC.SuppressFinalize(this)
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Cleansup up managed objects
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
                if (_Stream != null)
                {
                    _Stream.Dispose();
                    _Stream = null;
                }

                // free managed resources
                if (_Result != null)
                {
                    _Result.Dispose();
                    _Result = null;
                }
            }
        }

        # endregion
    }
}
