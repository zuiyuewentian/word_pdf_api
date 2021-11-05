using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.Extensions.Logging;
using Word;

namespace wordHelper.Controllers
{
    /// <summary>
    /// word转pdf
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {
        private readonly ILogger<WordController> _logger;


        private Application _word;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="logger"></param>
        public WordController(ILogger<WordController> logger)
        {
            _word = new Application();
            _logger = logger;
        }

        /// <summary>
        /// 获取
        /// </summary>
        /// <param name="formFile">word文档</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> Get(IFormFile formFile)
        {

            string fileDirectoryPath = AppContext.BaseDirectory + "/tempFiles/";
            try
            {
                if (formFile.Length > 0)
                {
                    var filePath = Path.Combine(fileDirectoryPath, formFile.FileName);
                    string extension = Path.GetExtension(filePath);

                    if (extension != ".doc" && extension != ".docx")
                    {
                        JsonResult res1 = new JsonResult(new { msg = "请传入word文件" });
                        return res1;
                    }

                    if (!Directory.Exists(fileDirectoryPath))
                    {
                        Directory.CreateDirectory(fileDirectoryPath);
                    }
                    string pdfPath = "";


                    var filePath2 = Path.Combine(fileDirectoryPath, Guid.NewGuid() + extension);

                    using (var stream = System.IO.File.Create(filePath2))
                    {
                        await formFile.CopyToAsync(stream);
                    }

                    if (extension == ".doc")
                    {
                        pdfPath = filePath.Replace(".doc", ".pdf");
                    }
                    if (extension == ".docx")
                    {
                        pdfPath = filePath.Replace(".docx", ".pdf");
                    }
                    ConvertToPdf(filePath2, pdfPath);

                    var fileName = Path.GetFileName(pdfPath);
                    var mimeType = "application/octet-stream";

                    var bys = System.IO.File.ReadAllBytes(pdfPath);
                    Stream mstream = new MemoryStream(bys);
                    mstream.Seek(0, SeekOrigin.Begin);

                    return new FileStreamResult(mstream, mimeType)
                    {
                        FileDownloadName = fileName
                    };

                }

                JsonResult res2 = new JsonResult(new { msg = "传入文件错误" });
                return res2;
            }
            catch (Exception ex)
            {
                JsonResult res3 = new JsonResult(new { msg = "异常" + ex.Message });
                return res3;
                // return "error" + ex.Message;
            }
            finally
            {
                Clearfiles(fileDirectoryPath);
            }
        }



        private void Clearfiles(string fileDirectoryPath)
        {
            try
            {
                var files = Directory.GetFiles(fileDirectoryPath);
                foreach (var item in files)
                {
                    System.IO.File.Delete(item);
                }
            }
            catch (Exception)
            {
            }
        }


        private bool ConvertToPdf(object sourcePath, object targetPath)
        {
            try
            {
                var result = true;
                Document doc = _word.Documents.Open(sourcePath, Visible: false);
                doc.ExportAsFixedFormat(targetPath.ToString(), WdExportFormat.wdExportFormatPDF);
                doc.Close();
                _word.Quit();
                return result;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}
