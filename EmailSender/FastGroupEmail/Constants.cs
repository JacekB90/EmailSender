using System;

namespace FastGroupEmail
{
    class Constants
    {
        public static string pathToDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public const string countMessage = " message has been sent.";
        public const string countMessagesPlural = " messages have been sent.";
        public const string confirmation = "Confirmation";
        public const string noMatchesStatement = "The program has not found a matching. ";
        public const string exitAsking = "Do you want exit?";
        public const string fillUp = "Please fill out all required fields.";
        public const string emptyPath = "Path(s) cannot be empty.";
        public const string emptySubject = "There is not an email subject.";
        public const string emptyMessage = "There is not an email message.";


    }
}
