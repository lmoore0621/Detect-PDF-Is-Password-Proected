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
            var wordFile = @"\\dwshome-a.homeoffice.wal-mart.com\dwsuserdata$\vn5007u\Desktop\ChangeRequests\CHG0504240_Validation Plan_QMS.docx";
            var wordFileProtected = @"\\dwshome-a.homeoffice.wal-mart.com\dwsuserdata$\vn5007u\Desktop\ChangeRequests\CHG0454645_Validation Plan.docx";
            var pdfFileProtected = @"C:\Toolbox\Source\Repos\Detect-PDF-Is-Password-Proected-master\Detect-PDF-Is-Password-Proected-master\DetectPassword\bin\776313.pdf";
            var pdfFileNotProtected = @"\\dwshome-a.homeoffice.wal-mart.com\dwsuserdata$\vn5007u\Desktop\Notes\QMS Database.pdf";

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
