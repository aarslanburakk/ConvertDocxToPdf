using Microsoft.Office.Interop.Word;
using System.IO.Compression;
using System.Runtime.InteropServices;

namespace ConvertDocx.Service.ConvertService
{
    public class ConvertFileService : IConvertFileService
    {
        /// <summary>
        /// This function does Convert one DOCX to one  PDF
        /// </summary>
        /// <param name="file"></param>
        /// <returns> Byte pdf array</returns>
        public byte[] ConvertFile(IFormFile file)
        {
            string uniqueFileName = Guid.NewGuid().ToString();
            string docxFilePath = Path.Combine(Path.GetTempPath(), uniqueFileName + ".docx");
            string pdfFilePath = Path.Combine(Path.GetTempPath(), uniqueFileName + ".pdf");

            // Save the uploaded file to a temporary location
            using (var stream = new FileStream(docxFilePath, FileMode.Create))
            {
                file.CopyTo(stream);
            }

            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(docxFilePath);
            doc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);
            doc.Close();
            wordApp.Quit();

            // Read the PDF file into a byte array
            byte[] pdfBytes = System.IO.File.ReadAllBytes(pdfFilePath);

            // Clean up temporary files
            System.IO.File.Delete(docxFilePath);
            System.IO.File.Delete(pdfFilePath);

            // Return the PDF file for download
            return pdfBytes;
        }
        /// <summary>
        /// This function does convert many DOCX to many PDF
        /// </summary>
        /// <param name="files"></param>
        /// <returns>MemoryStream zip file</returns>
        public MemoryStream ConvertZipFile(List<IFormFile> files)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var file in files)
                    {

                        string uniqueFileName = Guid.NewGuid().ToString();
                        string docxFilePath = Path.Combine(Path.GetTempPath(), uniqueFileName + ".docx");
                        string pdfFilePath = Path.Combine(Path.GetTempPath(), uniqueFileName + ".pdf");

                        using (var stream = new FileStream(docxFilePath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                        }

                        Application wordApp = new Application();
                        Document doc = wordApp.Documents.Open(docxFilePath);
                        doc.SaveAs2(pdfFilePath, WdSaveFormat.wdFormatPDF);
                        doc.Close();
                        wordApp.Quit();

                        byte[] pdfBytes = System.IO.File.ReadAllBytes(pdfFilePath);

                        // Add the converted PDF to the zip archive
                        var zipEntry = archive.CreateEntry(file.FileName.Replace(".docx", ".pdf"));
                        using (var zipStream = zipEntry.Open())
                        {
                            zipStream.Write(pdfBytes, 0, pdfBytes.Length);
                        }

                        // Clean up temporary files
                        System.IO.File.Delete(docxFilePath);
                        System.IO.File.Delete(pdfFilePath);
                    }
                }
                return memoryStream;
            }
        }
    }
}

