using ClosedXML.Excel;

namespace PracticTask3.Requests
{
    internal class ClientNameChangeRequest : IRequest
    {
        void IRequest.ExecuteRequest(XLWorkbook CurrentWorkbook) 
        {
            Console.Write("Введите название организации: ");
            string ?OrganisationName = Console.ReadLine();
            Console.Write("Введите новое ФИО контактного лица организации: ");
            string ?NewClientName = Console.ReadLine();

            
            bool OrganisationNameFound = false;

            var WorksheetClients = CurrentWorkbook.Worksheet("Клиенты");
            var OrganisationNamesRange = WorksheetClients.Range(WorksheetClients.Cell(2,"B"),WorksheetClients.Cell(WorksheetClients.RowCount(), "B"));

            foreach (var OrgName in OrganisationNamesRange.CellsUsed())
            {
                if (OrgName.Value.ToString() == OrganisationName)
                {
                    Console.WriteLine("Выполняю запрос на изменение ФИО контактного лица запрошенной организации...\n" +
                                      "==========================================================================================");
                    OrganisationNameFound = true;

                    var OrganisationNameRowNubmer = OrgName.WorksheetRow().RowNumber();
                    var OrganisationNameColumnNubmer = OrgName.WorksheetColumn().ColumnNumber();

                    var ClientNameCell = WorksheetClients.Cell(OrganisationNameRowNubmer, OrganisationNameColumnNubmer + 2);

                    ClientNameCell.Value = NewClientName;

                    try
                    {
                        xlsxFileManager.SaveFile(CurrentWorkbook);
                    }

                    catch (System.IO.IOException)
                    {
                        Console.WriteLine("Ошибка: Изменения не удалось сохранить в файл! Если файл используется в каком-то другом процессе, остановите его и попробуйте снова.\n" +
                            "Нажмите на любую клавишу, чтобы продолжить...");
                        Console.ReadLine();
                        break;
                    }

                    Console.WriteLine($">>Наименование запрошенной организации: {OrgName.Value.ToString()}\n" +
                                      $">>Новое контактное лицо запрошенной организации: {ClientNameCell.Value.ToString()}\n");
                    break;
                }
            }
            if (!OrganisationNameFound)
            {
                Console.WriteLine("==========================================================================================\n" +
                                  ">>Организации с таким названием нет в листе Клиенты.");
            }
            Console.WriteLine("Нажмите на любую клавишу, чтобы вернуться к списку доступных команд...");
            Console.ReadLine();
            Console.Clear();
        }
    }
}
