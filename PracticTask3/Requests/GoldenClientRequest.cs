using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;

namespace PracticTask3.Requests
{
    internal class GoldenClientRequest : IRequest
    {
        public int Id => 3;
        void IRequest.ExecuteRequest(XLWorkbook CurrentWorkbook) 
        {
            Console.WriteLine("Введите год и месяц в числовом формате за которые будет выведен клиент с наибольшим кол-вом заказов:");
            Console.Write("Год: ");
            string YearCheck = Console.ReadLine();

            int Year = 0;
            bool YearIsInt = true;

            try { Year = int.Parse(YearCheck); }
            catch(Exception) 
            {
                YearIsInt = false;
            }
 
            Console.Write("Месяц: ");
            string MonthCheck = (Console.ReadLine());
            
            int Month = 0;  
            bool MonthIsInt = true;

            try { Month = int.Parse(MonthCheck); }
            catch (Exception)
            {
                MonthIsInt = false;
            }

            if (YearIsInt && MonthIsInt)
            {
                var StartDate = new DateTime(Year, Month, 1);
                var EndDate = new DateTime(Year, Month, DateTime.DaysInMonth(Year, Month));

                bool DateExists = false;

                var WorksheetOrders = CurrentWorkbook.Worksheet("Заявки");
                var OrderDatesRange = WorksheetOrders.Range(WorksheetOrders.Cell(2, "F"), WorksheetOrders.Cell(WorksheetOrders.RowCount(), "F"));

                Dictionary<string, int> ClientCodes = new Dictionary<string, int>();

                foreach (var OrderDate in OrderDatesRange.CellsUsed())
                {
                    var CellsDate = OrderDate.GetDateTime();
                    if (CellsDate >= StartDate && CellsDate <= EndDate)
                    {
                        DateExists = true;
                        var CellsDateRow = OrderDate.WorksheetRow().RowNumber();
                        var CellsDateColumn = OrderDate.WorksheetColumn().ColumnNumber();

                        var ClientCodeCell = WorksheetOrders.Cell(CellsDateRow, CellsDateColumn - 3);
                        var ClientCodeValue = ClientCodeCell.Value.ToString();

                        if (ClientCodes.ContainsKey(ClientCodeValue))
                        {
                            ClientCodes[ClientCodeValue]++;
                        }
                        else
                        {
                            ClientCodes[ClientCodeValue] = 1;
                        }
                    }
                }
                string? GoldenClientCode = null;
                int ClientCodeCount = 0;

                foreach (var CheckClientCode in ClientCodes)
                {
                    if (CheckClientCode.Value > ClientCodeCount)
                    {
                        ClientCodeCount = CheckClientCode.Value;
                        GoldenClientCode = CheckClientCode.Key;
                    }
                }
                if (DateExists)
                {
                    Console.WriteLine("==========================================================================================\n" +
                                     $">>Код клиента с наибольшим кол-вом заказов за запрошенные даты: {GoldenClientCode}\n");
                }

                if (!DateExists)
                {
                    Console.WriteLine("==========================================================================================\n" +
                                      ">>Такой даты размещения заказа нет в листе Заявки.");
                }
                

            }
            else 
            {
                Console.WriteLine("==========================================================================================\n" +
                                  ">>Вы ввели данные в неверном формате.");
            }
            Console.WriteLine("Нажмите на любую клавишу, чтобы вернуться к списку доступных команд...");
            Console.ReadLine();
            Console.Clear();
        }
    }
}
