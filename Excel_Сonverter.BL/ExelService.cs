using OfficeOpenXml;

namespace Excel_Сonverter.BL
{
    public class ExelService
    {


        /// <summary>
        /// получим лист с данными из файла
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
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

        /// <summary>
        /// коды  получим отделно
        /// </summary>
        /// <param name="nameCity"></param>
        /// <param name="cityWorkbook"></param>
        /// <returns></returns>
        private string GetCodeCity(string nameCity, ExcelWorksheet cityWorkbook)
        {
            for (int i = 1; i<= cityWorkbook.Dimension.Rows; i++)
            {
                var r = cityWorkbook.Cells[i, 3].Text;
                if(r == nameCity)
                    return cityWorkbook.Cells[i, 2].Text;
            }
            return "код не найден";
        }


        /// <summary>
        /// получим цены 
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="row"></param>
        /// <param name="colCount"></param>
        /// <returns></returns>
        private List<PriceModel> GetPrices(ExcelRange cells, int row, int colCount)
        {
            List<PriceModel> models = new List<PriceModel>();
          
            int  x = 0;
          
            for (int i = 5; i < colCount; i+=3) // может быть  и больше  5 цен
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
     

        /// <summary>
        /// Сохранений  данных в  поток 
        /// </summary>
        /// <param name="models"></param>
        /// <returns></returns>
        public Stream SaveReports (List<ModelReport> models )
        {
            var memoryStream = new MemoryStream();

            using (var package = new ExcelPackage(memoryStream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Отчет");

                //ширина  
                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 30;
                worksheet.Column(3).Width = 30;
                worksheet.Column(4).Width = 30;
                worksheet.Column(5).Width = 30;
                worksheet.Column(6).Width = 30;
                worksheet.Column(7).Width = 40;
                worksheet.Column(8).Width = 50;
                
                //шапка
                worksheet.Cells[1, 1].Value = "№ п/п"; 
                worksheet.Cells[1, 2].Value = "Код города";     
                worksheet.Cells[1, 3].Value = "Наименование города"; 
                worksheet.Cells[1, 4].Value = "Всего Кол-во позиций в текущем городе"; 
                worksheet.Cells[1, 5].Value = "Кол-во позиций с отметкой «ЖЦ»"; 
                worksheet.Cells[1, 6].Value = "Кол-во позиций с отметкой «Т»"; 
                worksheet.Cells[1, 7].Value = "Кол-во пустых позиций, без цены, в текущем городе"; 
                worksheet.Cells[1, 8].Value = "Процент заполнения ценами в текущем городе"; 

                int r = 2; 

                //данные
                foreach (var item in models)
                {
                    worksheet.Cells[r, 1].Value = item.Position;   
                    worksheet.Cells[r, 2].Value = item.CodeCity;   
                    worksheet.Cells[r, 3].Value = item.NameCity;   
                    worksheet.Cells[r, 4].Value = item.CountProduct;  
                    worksheet.Cells[r, 5].Value = item.CountChZ; 
                    worksheet.Cells[r, 6].Value = item.CountT;   
                    worksheet.Cells[r, 7].Value = item.CountEmpty;   
                    worksheet.Cells[r, 8].Value = item.fillPercentage;   
                    r++; 
                }
               
                package.Save();
            }

            memoryStream.Position = 0;

            return memoryStream;
        }
    }
}

