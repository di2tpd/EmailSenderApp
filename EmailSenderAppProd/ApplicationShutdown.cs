using System;

namespace EmailSenderApp
{
    class ApplicationShutdown
    {
        private static int sendmailProcessesCount = 0;
        private static object lockObject = new object();

        public static void IncrementSendmailProcess()
        {
            lock (lockObject)
            {
                sendmailProcessesCount++;
            }
        }

        public static void DecrementSendmailProcess()
        {
            lock (lockObject)
            {
                sendmailProcessesCount--;
                if (sendmailProcessesCount == 0)
                {
                    CloseApplication();
                }
            }
        }

        public static void CloseApplication()
        {
            // Close the application here
            Environment.Exit(0);
        }
    }
}
