using OfficeOpenXml;
using ReadExcelEpPlus.Enums;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ReadExcelEpPlus
{
    public class ReadExcel
    {
        public async Task<List<Car>> Read(ExcelPackage file)
        {
            var cars = new List<Car>();
            var worksheet = (int)ExcelWorksheetEnum.cars;

            using (var package = file)
            {
                var totalRows = package.Workbook.Worksheets[worksheet].Dimension?.Rows;

                for (int j = 1; j <= totalRows.Value; j++)
                {
                    FillCar(package, worksheet, j, cars);
                }
            }

            return cars;
        }

        private void FillCar(ExcelPackage excelPackage, int i, int j, List<Car> cars)
        {
            var car = new Car();

            car.Name = excelPackage.Workbook.Worksheets[i].Cells[j, (int)CarExcelColumnEnum.Name].Value.ToString();
            car.Brand = excelPackage.Workbook.Worksheets[i].Cells[j, (int)CarExcelColumnEnum.Brand].Value.ToString();
            car.Price = excelPackage.Workbook.Worksheets[i].Cells[j, (int)CarExcelColumnEnum.Price].Value.ToString();

            cars.Add(car);
        }
    }
}
