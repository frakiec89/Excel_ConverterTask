using OfficeOpenXml;

namespace Excel_Сonverter.BL
{
    public class ExelService
    {


        public List<ModelExel> GetModelExels(MemoryStream stream)
        {
            List<ModelExel> models = new List<ModelExel>();

            // Открытие и чтение Excel файла
            using (var package = new ExcelPackage(stream))
            {
                // на  основе  моего  кода  https://github.com/frakiec89/Parsing_RSO_To_GSK.WebAPI/blob/master/Parsing_RSO_To_GSK.WebAPI.ParseExel/ExelService.cs 
                ExcelPackage.License.SetNonCommercialOrganization("<My Noncommercial organization>");
                ExcelPackage.License.SetNonCommercialPersonal("<My Name>");

                // Получаем первый лист

                var cityWorkbook = package.Workbook.Worksheets.SingleOrDefault(x=>x.Name.StartsWith("Регионы продаж"));

                foreach (var worksheet in package.Workbook.Worksheets.Where(x=>x.Name.StartsWith("Регионы продаж")==false))
                {
                    // Получаем количество строк и столбцов
                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;
                    
                    string cityName = worksheet.Name;
                    string cityCode = GetCodeCity(cityName, cityWorkbook);
                    int position = 0; 
                    // Чтение данных
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var idStr = worksheet.Cells[row, 2].Text;

                        if (string.IsNullOrWhiteSpace(idStr) || idStr.ToLower().Contains("id"))
                            continue;

                        position++;

                        var m = new ModelExel();
                        m.NameProduct = worksheet.Cells[row, 3].Text.Trim();
                        m.ID = worksheet.Cells[row, 3].Text.Trim();
                        m.Unit = worksheet.Cells[row, 4].Text.Trim();
                        m.Position = position;
                        m.Prices = GetPrices(worksheet.Cells, row , colCount);
                        m.NameCity = cityName;
                        m.CodeCity = cityCode;   
                        models.Add(m);
                    }
                }
            }
            return models;
        }

        private string GetCodeCity(string nameCity, ExcelWorksheet cityWorkbook)
        {
            for (int i = 1; i< cityWorkbook.Dimension.Rows; i++)
            {
                var r = cityWorkbook.Cells[i, 3].Text;
                if(r == nameCity)
                    return cityWorkbook.Cells[i, 2].Text;
            }
            return "код не найден";
        }

        private List<PriceModel> GetPrices(ExcelRange cells, int row, int colCount)
        {
            List<PriceModel> models = new List<PriceModel>();
          
            int  x = 0;
          
            for (int i = 5; i < colCount; i+=3)
            {
                x++;
            
                PriceModel model = new PriceModel();

                var r = cells [ row , i].Text.Trim();

                if (double.TryParse(r, out double v))
                    model.Value = v;
                else
                    model.Value = 0;
            
                model.Id = x;
                model.PriceIndicator = cells[row , i+1].Text.Trim();
                model.Supplier = cells[row , i+2].Text.Trim();
                models.Add(model);
            }

            return models;
        }
        
    }
}

