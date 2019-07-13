using System;
using System.IO;

namespace FastGroupEmail
{
    class FileReader
    {
        public string[] ReadMailList(string filename)
        {
            string[] fileContent;
            try
            {
                fileContent = File.ReadAllLines(filename);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return fileContent;
        }
        public string[] ReadAttachments(string filename)
        {
            string[] fileContentAttachment;
            try
            {
                fileContentAttachment = Directory.GetFiles(filename);
                
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return fileContentAttachment;
        }
    }
}
