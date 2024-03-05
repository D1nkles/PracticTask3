using ClosedXML.Excel;

namespace PracticTask3
{
    static class xlsxFileManager
    {
        public static XLWorkbook GetFile(string PATH) 
        {
            XLWorkbook CurrentWorkbook = new XLWorkbook(PATH);
            return CurrentWorkbook;
        }

        public static void SaveFile(XLWorkbook CurrentWorkbook) 
        {
                CurrentWorkbook.Save();
        }
        
    }
    
}
