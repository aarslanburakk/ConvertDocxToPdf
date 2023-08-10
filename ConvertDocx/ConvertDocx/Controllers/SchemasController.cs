using ConvertDocx.Service.ConvertService;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;

namespace ConvertDocx.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SchemasController : ControllerBase
    {
        private readonly IConvertFileService _convertFileService;
        public SchemasController(IConvertFileService convertFileService)
        {
            _convertFileService = convertFileService;
        }

        [HttpPost]
        [Route("OneDocxFileInput")]
        public IActionResult ConvertPdf(IFormFile file)
        {
            if (file == null || file.Length <= 0)
            {
                return BadRequest("No file was uploaded.");
            }
            if (Path.GetExtension(file.FileName) != ".docx")
            {
                return BadRequest("Only .docx files are supported.");
            }
            // Return the PDF file for download
            return File(_convertFileService.ConvertFile(file), "application/pdf", $"{file.FileName}.pdf");

        }

        [HttpPost]
        [Route("MultipleDocxFileInput")]
        public IActionResult ConvertAndZipPdf(List<IFormFile> files)
        {

            if (files == null || files.Count == 0)
            {
                return BadRequest("No files were uploaded.");
            }


            // Return the zip file for download
            return File(_convertFileService.ConvertZipFile(files).ToArray(), "application/zip", "ConvertedFiles.zip");
        }




    }




}

