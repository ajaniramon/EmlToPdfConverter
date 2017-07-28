using EAGetMail;
using MailToPdfConverter.Util;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace MailToPdfConverter.Logic
{
    public class MailLogic
    {
        public static string ConvertToZippedPfdWithAttachments(string base64Eml)
        {
            // Carpeta donde se va a almacenar temporalmente el mail
            string tempFolder = "C:\\tempMail\\";

            // Ruta destino del archivo HTML generado
            string htmlName = System.IO.Path.Combine(tempFolder, "message.html");

            // Se crea el directorio donde se va a almacenar todo
            if (!System.IO.Directory.Exists(tempFolder))
            {
                System.IO.Directory.CreateDirectory(tempFolder);
            }

            // Se obtienen los archivos que hay en el directorio
            IEnumerable<string> totalFilesInFolder =
                System.IO.Directory.EnumerateFiles(tempFolder);

            // Se limpia el directorio para que no haya nada excepto nuestro mensaje
            if (totalFilesInFolder.Count() > 0)
            {
                foreach (var file in totalFilesInFolder)
                {
                    System.IO.File.Delete(file);
                }
            }

            // Se convierte a array de bytes
            byte[] fileBytes = Convert.FromBase64String(base64Eml);

            Mail mail = new Mail("TryIt");

            mail.Load(fileBytes);

            // Parse html body
            string html = mail.HtmlBody;
            StringBuilder hdr = new StringBuilder();

            // Parse sender
            hdr.Append("<font face=\"Courier New,Arial\" size=2>");
            hdr.Append("<b>From:</b> " + MailHelper._FormatHtmlTag(mail.From.ToString()) + "<br>");

            // Parse to
            MailAddress[] addrs = mail.To;
            int count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>To:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(MailHelper._FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }

            // Parse cc
            addrs = mail.Cc;

            count = addrs.Length;
            if (count > 0)
            {
                hdr.Append("<b>Cc:</b> ");
                for (int i = 0; i < count; i++)
                {
                    hdr.Append(MailHelper._FormatHtmlTag(addrs[i].ToString()));
                    if (i < count - 1)
                    {
                        hdr.Append(";");
                    }
                }
                hdr.Append("<br>");
            }

            hdr.Append(String.Format("<b>Subject:</b>{0}<br>\r\n",
               MailHelper._FormatHtmlTag(mail.Subject)));

            // Parse attachments and save to local folder
            Attachment[] atts = mail.Attachments;
            count = atts.Length;
            if (count > 0)
            {
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);

                hdr.Append("<b>Attachments:</b>");
                for (int i = 0; i < count; i++)
                {
                    Attachment att = atts[i];

                    // this attachment is in OUTLOOK RTF format, decode it here.
                    if (String.Compare(att.Name, "winmail.dat") == 0)
                    {
                        Attachment[] tatts = null;
                        try
                        {
                            tatts = Mail.ParseTNEF(att.Content, true);
                        }
                        catch (Exception ep)
                        {
                            Console.WriteLine(ep.Message);
                            continue;
                        }

                        int y = tatts.Length;
                        for (int x = 0; x < y; x++)
                        {
                            Attachment tatt = tatts[x];
                            string tattname = String.Format("{0}/{1}", tempFolder, tatt.Name);
                            tatt.SaveAs(tattname, true);
                            hdr.Append(
                            String.Format("<a target=\"_blank\">{1}</a> ",
                                tattname, tatt.Name));
                        }
                        continue;
                    }

                    string attname = String.Format("{0}/{1}", tempFolder, att.Name);
                    att.SaveAs(attname, true);
                    hdr.Append(String.Format("<a  target=\"_blank\">{1}</a> ",
                            attname, att.Name));
                    if (att.ContentID.Length > 0)
                    {
                        string base64content = Convert.ToBase64String(System.IO.File.ReadAllBytes(attname));
                        // Show embedded images.
                        html = html.Replace("cid:" + att.ContentID, "data:image/png;base64," + base64content);
                    }
                    else if (String.Compare(att.ContentType, 0, "image/", 0,
                                "image/".Length, true) == 0)
                    {
                        string base64content = Convert.ToBase64String(System.IO.File.ReadAllBytes(attname));
                        // show attached images.
                        html = html + String.Format("<hr><img src=\"{0},{1}\">", "data:image/png;base64,", base64content);
                    }
                }
            }

            Regex reg = new Regex("(<meta[^>]*charset[ \t]*=[ \t\"]*)([^<> \r\n\"]*)",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
            html = reg.Replace(html, "$1utf-8");
            if (!reg.IsMatch(html))
            {
                hdr.Insert(0,
                    "<meta HTTP-EQUIV=\"Content-Type\" Content=\"text-html; charset=utf-8\">");
            }

            // write html to file
            html = hdr.ToString().Replace("(Trial Version)", "") + "<hr>" + html;
            FileStream fs = new FileStream(htmlName, FileMode.Create,
                FileAccess.Write, FileShare.None);

            byte[] fileData = System.Text.UTF8Encoding.UTF8.GetBytes(html);
            fs.Write(fileData, 0, fileData.Length);
            fs.Close();
            mail.Clear();


            // instantiate the html to pdf converter
            HtmlToPdf converter = new HtmlToPdf();

            // convert the url to pdf
            PdfDocument doc = converter.ConvertHtmlString(html);

            // save pdf document
            doc.Save(System.IO.Path.Combine(tempFolder, "message.pdf"));

            // close pdf document
            doc.Close();

            string zippedMailsFolder = "C:\\zippedMails\\";

            // Una vez tenemos nuestra carpeta lista la vamos a comprimir
            string randomFolderName = System.IO.Path.Combine(zippedMailsFolder,"mail_zipped_" + DateTime.Now.Ticks + ".zip");

            ZipFile.CreateFromDirectory(tempFolder, randomFolderName);

            string result = string.Empty;
            if (System.IO.File.Exists(randomFolderName))
            {
                result = Convert.ToBase64String(System.IO.File.ReadAllBytes(randomFolderName));
            }
            return result;
        }
    }
}