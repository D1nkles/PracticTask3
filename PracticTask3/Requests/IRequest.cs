using ClosedXML.Excel;

namespace PracticTask3.Requests
{
    public interface IRequest
    {
        int Id { get; }
        void ExecuteRequest(XLWorkbook CurrentWorkbook) 
        {

        }
    }
}
