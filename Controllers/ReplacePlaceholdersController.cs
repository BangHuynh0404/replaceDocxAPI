using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ReplacePlaceholdersApi.Controllers
{
   [Route("test")]
   [ApiController]
   public class ReplacePlaceholdersController : ControllerBase
   {
      [HttpPost]
      public async Task<IActionResult> ReplacePlaceholders(IFormFile file)
      {
         try
         {
            using (var stream = new MemoryStream())
            {
               await file.CopyToAsync(stream);
               stream.Seek(0, SeekOrigin.Begin);

               using (var wordDocument = WordprocessingDocument.Open(stream, true))
               {
                  var mainPart = wordDocument.MainDocumentPart;
                  var text = mainPart.Document.Body.InnerText;

                  JObject data = JObject.Parse(System.IO.File.ReadAllText("data.json"));
                  Console.WriteLine(data);
                  foreach (var property in data.Properties())
                  {
                     text = text.Replace("{" + property.Name + "}", property.Value.ToString());
                  }

                  mainPart.Document.Body.RemoveAllChildren();
                  mainPart.Document.Body.Append(new Paragraph(new Run(new Text(text))));

                  wordDocument.Save();
                  Console.WriteLine("Placeholders replaced successfully with data from the JSON file.");
               }


               stream.Seek(0, SeekOrigin.Begin);
               return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "replaced.docx");
            }
         }
         catch (Exception ex)
         {
            return BadRequest(ex.Message);
         }
      }
   }
}
