using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TeamsAutoJoiner
{
    class Program
    {
        static Teams t = null;

        #region Trap application termination

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);
        [DllImport("Kernel32")]
        private static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);

        private delegate bool EventHandler(CtrlType sig);
        static EventHandler _handler;

        enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        static bool HTirggered = false;
        private static bool Handler(CtrlType sig)
        {
            if (HTirggered) return false;
            HTirggered = true;
            t.Close();
            Environment.Exit(-1);
            return true;
        }
        #endregion

        enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
        };

        [DllImport("Kernel32.dll")]
        private static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE state);
        static void Main(string[] args)
        {
            if(SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS | EXECUTION_STATE.ES_SYSTEM_REQUIRED) == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("SetThreadExecutionState failed.");
                return;
            }
            _handler += new EventHandler(Handler);
            SetConsoleCtrlHandler(_handler, true);
            try
            {
                t = new Teams();
                t.Work(args);

                Handler(CtrlType.CTRL_C_EVENT);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
