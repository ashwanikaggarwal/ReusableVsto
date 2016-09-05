using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Reusable.Threading
{
    public interface ITest
    {
        void Run();
    }

    public class ThreadArray
    {
        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern void Sleep(long dwMilliseconds);

        private int _Threads;
        private bool _CancelThread;
        private string _RunStamp;
        private string _Errors;

        private long _Sleep;
        private System.DateTime _ProcessStart;

        private System.DateTime _ProcessEnd;

        public string RunStamp
        {
            get { return _RunStamp; }
            set { _RunStamp = value; }
        }

        private Boolean _Completed = false;
        public Boolean Completed
        {
            get { return _Completed; }
        }
        
        public System.DateTime ProcessStart
        {
            get { return _ProcessStart; }
        }

        public System.DateTime ProcessEnd
        {
            get { return _ProcessEnd; }
        }

        public int Concurrency
        {
            get { return _Threads; }
            set { _Threads = value; }
        }

        public bool Cancel
        {
            get { return _CancelThread; }
            set { _CancelThread = value; }
        }

        public string ThreadError
        {
            get { return _Errors; }
        }

        public long SecondsBetweenThreads
        {
            get { return _Sleep; }
            set { _Sleep = value; }
        }

        private void SleepProcess(long lngSleep)
        {
            Sleep(lngSleep * 1000);
        }

        public void Execute(ITest[] Tests)
        {
            _ProcessStart = DateTime.Now;

            Thread[] threadList = new Thread[_Threads];

            try
            {
                _RunStamp = String.Format(DateTime.Now.ToString("yyyyMMddHHmmss"));
                for (int intCounter = 0; intCounter <= _Threads - 1; intCounter++)
                {
                    ThreadStart threadDelegate = new ThreadStart(Tests[intCounter].Run);
                    threadList[intCounter] = new Thread(threadDelegate);
                    threadList[intCounter].Start();
                    if (intCounter > 0)
                    {
                        SleepProcess(_Sleep);
                    }
                }

                foreach (Thread t in threadList)
                {
                    t.Join();
                }
            }
            catch (Exception ex)
            {
                _Errors = ex.Message.ToString();
            }
            finally
            {
                _ProcessEnd = DateTime.Now;

                for (Int32 counter = 0; counter < threadList.Length; counter++)
                {
                    threadList[counter] = null;
                }
            }

            _Completed = true;
        }

        public ThreadArray()
        {
            _RunStamp = "0";
            _Errors = "";
            _Sleep = 0;
        }
    }
}
