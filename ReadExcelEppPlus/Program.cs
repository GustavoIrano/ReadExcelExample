using OfficeOpenXml;
using ReadExcelEpPlus;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ReadExcelEppPlus
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var file = new ExcelPackage(new FileInfo("C:\\SEFIP\\cars.xlsx"));
            var cars = await new ReadExcel().Read(file);

            Console.WriteLine("Name - Brand - Price");
            Console.WriteLine();

            foreach (var car in cars)
            {
                Console.WriteLine(car.Name + " - " + car.Brand + " - " + car.Price);
            }

        }
    }
}
