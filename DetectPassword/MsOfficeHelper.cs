using iTextSharp.text.exceptions;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DetectPassword
{
    public static class MsOfficeHelper
    {
        public static bool IsPasswordProtected(string pdfFile)
        {
            var isProtected = false;
            var isPdf = IsPdf(pdfFile);

            if (pdfFile != null && isPdf)
            {
                try
                {
                    PdfReader pdfReader = new PdfReader(pdfFile);
                }
                catch (BadPasswordException)
                {
                    isProtected = true;
                }
            }
            return isProtected;
        }

        public static string PdfProtectionErrorMessage(bool pdfIsProtected)
        {
            var message = "";
            if (pdfIsProtected)
            {
                message = "This File is password Protected, please upload a pdf document that is not password protected.";
            }
            return message;
        }

        private static bool IsPdf(string path)
        {
            //ÐÏ.à¡±.á
            var pdfString = "%PDF-";
            var pdfBytes = Encoding.ASCII.GetBytes(pdfString);
            var len = pdfBytes.Length;
            var buf = new byte[len];
            var remaining = len;
            var pos = 0;
            using (var f = File.OpenRead(path))
            {
                while (remaining > 0)
                {
                    var amtRead = f.Read(buf, pos, remaining);
                    if (amtRead == 0) return false;
                    remaining -= amtRead;
                    pos += amtRead;
                }
            }
            return pdfBytes.SequenceEqual(buf);
        }
    }
}