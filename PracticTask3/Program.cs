using ClosedXML.Excel;
using PracticTask3;

class Program
{
    static void Main()
    {
        Console.WriteLine("Добро пожаловать в программу для взаимодействия с конкретным Excel файлом!!!\n" +
                          "==========================================================================================");
        Console.WriteLine("Для начала работы введите путь к папке с Excel файлом под названием «Практическое задание для кандидата»");
        string PATH = Console.ReadLine() + "\\Практическое задание для кандидата.xlsx";
        
        var CurrentWorkbook = new XLWorkbook();
        try 
        { 
        XLWorkbook wb = xlsxFileManager.GetFile(PATH);
        CurrentWorkbook = wb;
        }
        catch (System.IO.FileNotFoundException) 
        {
            Console.WriteLine("Ошибка: По данному пути файл не найден! Перезапустите приложение, чтобы попробовать снова.\n" +
                              "Нажмите любую клавишу, чтобы закрыть приложение...");
            Console.ReadLine();
            return;
        }
        catch (System.IO.IOException)
        {
            Console.WriteLine("Ошибка: Файл не удалось загрузить! Если файл используется в каком-то другом процессе, остановите его и перезапустите приложение.\n" +
                              "Нажмите любую клавишу, чтобы закрыть приложение...");
            Console.ReadLine();
            return;
        }

        RequestManager requestManager = new RequestManager();
        requestManager.RequestSelecter(CurrentWorkbook);
    }
}

