using ClosedXML.Excel;

namespace PracticTask3.Requests
{
    internal class ProductInfoRequest : IRequest
    {
        void IRequest.ExecuteRequest(XLWorkbook CurrentWorkbook) 
        {
            Console.Write("Введите наименование товара: ");
            string ?ProductName = Console.ReadLine();

            Console.WriteLine($"Информация о клиентах, заказавших {ProductName}:");
            var WorksheetProducts = CurrentWorkbook.Worksheet("Товары");
            var ProductNameRange = WorksheetProducts.Range(WorksheetProducts.Cell(2, "B"), WorksheetProducts.Cell(WorksheetProducts.RowCount(), "B"));

            var WorksheetClients = CurrentWorkbook.Worksheet("Клиенты");
            var ClientIdRange = WorksheetClients.Range(WorksheetClients.Cell(2, "A"), WorksheetClients.Cell(WorksheetClients.RowCount(), "A"));

            var WorksheetOrders = CurrentWorkbook.Worksheet("Заявки");
            var ProductIdRange = WorksheetOrders.Range(WorksheetOrders.Cell(2, "B"), WorksheetOrders.Cell(WorksheetOrders.RowCount(), "B"));

            bool ProductFound = false;
            bool ProductIdFound = false;
            bool ClientFound = false;

            foreach (var ProdName in ProductNameRange.CellsUsed())
            {  
                if (ProdName.Value.ToString() == ProductName)
                {
                    ProductFound = true;

                    int ProductNameRowNumber = ProdName.WorksheetRow().RowNumber();
                    int ProductNameColumnNumber = ProdName.WorksheetColumn().ColumnNumber();

                    var ProductId = WorksheetProducts.Cell(ProductNameRowNumber, ProductNameColumnNumber - 1).Value.ToString();
                    var ProductValue = WorksheetProducts.Cell(ProductNameRowNumber, ProductNameColumnNumber + 1).Value.ToString();
                    var ProductPrice = WorksheetProducts.Cell(ProductNameRowNumber, ProductNameColumnNumber + 2).Value.ToString();

                    foreach (var ProdCode in ProductIdRange.CellsUsed())
                    {   
                        if (ProdCode.Value.ToString() == ProductId)
                        {
                            ProductIdFound = true;

                            Console.WriteLine("==========================================================================================");

                            int IdRowNumber = ProdCode.WorksheetRow().RowNumber();
                            int IdColumnNumber = ProdCode.WorksheetColumn().ColumnNumber();

                            var OrderDate = WorksheetOrders.Cell(IdRowNumber, IdColumnNumber + 4).Value.ToString();
                            var OrderAmount = WorksheetOrders.Cell(IdRowNumber, IdColumnNumber + 3).Value.ToString();
                            var ClientId = WorksheetOrders.Cell(IdRowNumber, IdColumnNumber + 1).Value.ToString();

                            foreach (var ClientCode in ClientIdRange.CellsUsed())
                            { 
                                if (ClientCode.Value.ToString() == ClientId)
                                {
                                    ClientFound = true;
                                   
                                    int ClientIdRowNumber = ClientCode.WorksheetRow().RowNumber();
                                    int ClientIdColumnNumber = ClientCode.WorksheetColumn().ColumnNumber();

                                    var ClientName = WorksheetClients.Cell(ClientIdRowNumber, ClientIdColumnNumber + 3).Value.ToString();
                                    var ClientAdress = WorksheetClients.Cell(ClientIdRowNumber, ClientIdColumnNumber + 2).Value.ToString();
                                    var OranisationName = WorksheetClients.Cell(ClientIdRowNumber, ClientIdColumnNumber + 1).Value.ToString();

                                    Console.WriteLine($">>ФИО клиента: {ClientName}.\n" +
                                                      $">>Адрес клиента: {ClientAdress}.\n" +
                                                      $">>Название организации клиента: {OranisationName}.");

                                    Console.WriteLine($">>Кол-во товара в заказе: {OrderAmount}.\n" +
                                                      $">>Цена за один {ProductValue} товара: {ProductPrice} рублей.\n" +
                                                      $">>Цена всего заказа: {int.Parse(OrderAmount) * int.Parse(ProductPrice)} рублей.\n" +
                                                      $">>Дата заказа: {OrderDate}.");
                                    Console.WriteLine("==========================================================================================\n");
                                }
                            } 
                            if (!ClientFound)
                            {
                                Console.WriteLine("==========================================================================================\n" +
                                                  ">>Клиент, заказавший товар не найден в листе Клиенты, поэтому о нем нет никакой информации.");
                            }
                        }
                    }
                    if (!ProductIdFound)
                    {
                        Console.WriteLine("==========================================================================================\n" +
                                          ">>На данный продукт не оформляли заявок.");
                    }
                }
            }
            if (!ProductFound)
            {
                Console.WriteLine("==========================================================================================\n" +
                                  ">>Продукта с таким наименованием нет в листе Товары.");
            }
            Console.WriteLine("Нажмите на любую клавишу, чтобы вернуться к списку доступных команд...");
            Console.ReadLine();
            Console.Clear();
        }
    }
}
