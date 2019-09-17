using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DetectPassword
{
    class Program
    {
        static void Main(string[] args)
        {
            //776313
            //Resume
            var wordFile = @"c";
            var wordFileProtected = @"c";
            var pdfFileProtected = @"Cc";
            var pdfFileNotProtected = @"cc";

            var passwordProtectedPDF = false;
            var isPasswordProtectedDoc = false;

            FileInfo fi = new FileInfo(wordFile);
            var ext = fi.Extension.ToLower();

            if (ext == ".pdf")
            {
                passwordProtectedPDF = MsOfficeHelper.IsPasswordProtectedPdf(pdfFileProtected);
            }
            else if (ext == ".doc" || ext == ".docx")
            {
                isPasswordProtectedDoc = MsOfficeHelper.IsPasswordProtectedWordDoc(wordFile);
            }

            var errorMessage = MsOfficeHelper.ProtectionErrorMessage(passwordProtectedPDF, isPasswordProtectedDoc);
            Console.WriteLine(errorMessage);
            Console.ReadLine();
        }
    }
}
