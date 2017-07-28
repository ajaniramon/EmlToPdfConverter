using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MailToPdfConverter.Util
{
    public class MailHelper
    {
        public static string _FormatHtmlTag(string src)
        {
            src = src.Replace(">", "&gt;");
            src = src.Replace("<", "&lt;");
            return src;
        }

    }
}