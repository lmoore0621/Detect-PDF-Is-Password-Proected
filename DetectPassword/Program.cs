using System;
using System.Collections.Generic;
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
            string nonPdfFile = @"C:\Users\lmoor\OneDrive\Desktop\Resume.docx";
            string pdfFileProtected = @"C:\Users\lmoor\OneDrive\Desktop\776313.pdf";
            string pdfFileNotProtected = @"C:\Users\lmoor\OneDrive\Desktop\Resume.pdf";

            bool passwordProtectedPDF = MsOfficeHelper.IsPasswordProtected(pdfFileProtected);

            if (passwordProtectedPDF)
            {
                string pdfProtectionErrorMessage = MsOfficeHelper.PdfProtectionErrorMessage(passwordProtectedPDF);
                Console.WriteLine(pdfProtectionErrorMessage);
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(passwordProtectedPDF);
                Console.ReadLine();
            }
        }
    }
}
