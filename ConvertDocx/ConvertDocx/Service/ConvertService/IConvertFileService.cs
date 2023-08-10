namespace ConvertDocx.Service.ConvertService
{
    public interface IConvertFileService
    {
        public byte[] ConvertFile(IFormFile file);
        public MemoryStream ConvertZipFile(List<IFormFile> files);
        

    }
}
