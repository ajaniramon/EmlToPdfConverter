using EAGetMail;
using MailToPdfConverter.Logic;
using MailToPdfConverter.Util;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Http;
using Toolkit.Helpers;

namespace MailToPdfConverter.Controllers
{
    /// <summary>
    /// Controlador que convierte un fichero .eml a un paquete completo ZIP con un PDF y sus adjuntos
    /// </summary>
    public class MailController : ApiController
    {
        [HttpPost]
        public HttpResponseMessage ConvertToPdf([FromBody] dynamic data)
        {
            // Se recoge el base64 del fichero
            string base64 = data.emlBase64;

            // Variable para almacenar el resultado
            string result = string.Empty;

            try
            {
                // Se convierte el mail y se obtiene el Base64 del resultado
                result = MailLogic.ConvertToZippedPfdWithAttachments(base64);
            }
            catch (Exception ex)
            {
                ex.registerException<MailController>();

                return new HttpResponseMessage()
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    Content = new StringContent(ex.Message)
                };
            }

            // Se retorna el resultado
            return new HttpResponseMessage()
            {
                Content = new StringContent(result),
                StatusCode = HttpStatusCode.OK
            };
        }
    }
}
