using System.Diagnostics;

namespace ExcelAutomationToolLibrary
{
    public class Logger
    {
        private Logger()
        {
            Trace.Listeners.Add(new TextWriterTraceListener("ExcelAutomationTool.log", "LoggerListener"));
        }

        //Return class instatnce.
        public static Logger Instance { get => new Logger(); }

        /// <summary>
        /// Log msg to log file.
        /// </summary>
        /// <param name="message">String type contains message to log.</param>
        public void LogMessage(string message)
        {
            Trace.Indent();
            Trace.TraceInformation(message);
            Trace.TraceError(message); // Testing... Output to check in log file.
            Trace.Unindent();
            Trace.Flush();
        }
    }
}
