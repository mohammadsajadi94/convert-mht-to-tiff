using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using Aspose.Email;
using NReco.ImageGenerator;

namespace convert_mht_to_tiff
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FileStream fileMht = new FileStream("enter path file here", FileMode.Open);
            byte[] bfile = new byte[fileMht.Length];
            fileMht.Read(bfile, 0, Convert.ToInt32(bfile.Length));
            byte[] TiffFile = ConvertMhtToTiff(bfile);
            fileMht.Close();
        }
        public static byte[] ConvertMhtToTiff(byte[] byteArray)
        {
            MemoryStream st = new MemoryStream(byteArray);
            var message = Aspose.Email.MailMessage.Load(st);
            string tempPath = System.IO.Path.GetTempPath() + DateTime.Now.Year.ToString() + "\\";
            if (System.IO.Directory.Exists(tempPath))
            {
                DirectoryInfo dir = new System.IO.DirectoryInfo(tempPath);
                dir.Delete(true);
            }
            System.IO.Directory.CreateDirectory(tempPath);
            string fileName = tempPath + "output.html";

            message.Save(fileName, Aspose.Email.SaveOptions.DefaultHtml);
            StreamReader reader = File.OpenText(fileName);
            string source = reader.ReadToEnd();
            reader.Close();
            File.Delete(fileName);
            string rep = @"<br><center>This is an evaluation copy of Aspose.Email for .NET</center><br><a href=""http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx""><center>View EULA Online</center></a><hr>";
            System.Text.RegularExpressions.Regex htmlregex = new System.Text.RegularExpressions.Regex(rep);
            string source1 = htmlregex.Replace(source, string.Empty);
            var conv = new NReco.ImageGenerator.HtmlToImageConverter();
            var image = conv.GenerateImage(source1, NReco.ImageGenerator.ImageFormat.Jpeg);
            string tiff_file = tempPath + "output.tiff";
            MemoryStream ms = new MemoryStream(image);
            Bitmap img = new Bitmap(ms);
            img.Save(tiff_file, System.Drawing.Imaging.ImageFormat.Tiff);
            byte[] byte_file = File.ReadAllBytes(tiff_file);
            img.Dispose();
            return byte_file;
        }
    }
}
