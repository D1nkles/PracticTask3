using ClosedXML.Excel;
using PracticTask3.Requests;

namespace PracticTask3
{
    internal class RequestManager
    {
        public void RequestSelecter(XLWorkbook CurrentWorkbook) 
        {
            var RequestStorage = new Dictionary<string, IRequest> 
            {
                ["1"] = new ProductInfoRequest() ,
                ["2"] = new ClientNameChangeRequest(),
                ["3"] = new GoldenClientRequest(),
            };

            while (true) 
            {
                Console.WriteLine("==========================================================================================");
                Console.WriteLine("Список доступных команд: \n" +
                                  "1. По наименованию товара вывести информацию о клиентах, заказавших этот товар.\n" +
                                  "2. Изменить контактное лицо клиента.\n" +
                                  "3. Определить золотого клиента за указанный год, месяц.\n" +
                                  "4. Закрыть программу.");
                Console.WriteLine("==========================================================================================\n");

                Console.Write("Введите номер команды, которую вы хотите исполнить: ");
                string UserInput = Console.ReadLine();

                if (RequestStorage.ContainsKey(UserInput)) 
                {
                    var Request = RequestStorage[UserInput];
                    Request.ExecuteRequest(CurrentWorkbook);
                }

                if (UserInput == "4")
                {
                    Environment.Exit(0);
                }

                if (!RequestStorage.ContainsKey(UserInput)) 
                {
                    Console.WriteLine("Команды с таким номером нет в списке доступных команд! Нажмите на любую клавишу, чтобы поробовать снова...");
                    Console.ReadLine();
                    Console.Clear();
                    continue;
                } 
            }
        }
    }
}
