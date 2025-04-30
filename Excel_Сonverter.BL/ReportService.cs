

namespace Excel_Сonverter.BL
{
    public class ReportService
    {
        public List<ModelReport> GetModelReports (List<ModelExel> models)
        {
            List<string> cities = models.Select(x=>x.CodeCity).Distinct().ToList();  // список  городов 
            List<ModelReport> modelReports = new List<ModelReport>();

            int x = 0;  
            foreach (var codeCity in cities)
            {
                // todo переписать  на linq с группировками

                var cityReport = models.Where(x => x.CodeCity == codeCity); // продукты  в  однм город 
                List<PriceModel> prices = GetListPrices(cityReport); 

                x++; // позиция

                var m = new ModelReport();

                m.NameCity = cityReport.FirstOrDefault().NameCity;
                m.CodeCity = codeCity;
                m.Position = x;
                m.CountProduct = cityReport.Count();
                m.CountChZ = prices.Where(x=>x.PriceIndicator.ToLower().StartsWith("жц")).Count();
                m.CountT =  prices.Where(x => x.PriceIndicator.ToLower().StartsWith("т")).Count();
                m.CountEmpty = prices.Where(x => string.IsNullOrEmpty(x.PriceIndicator)).Count();
                m.fillPercentage = ((m.CountChZ + m.CountT) / 5.0) *100;// todo поменять 5  ; формула странная
                modelReports.Add(m);
            }
            return modelReports;
        }


        /// <summary>
        /// Все цены в городе  в лист
        /// </summary>
        /// <param name="cityReport"></param>
        /// <returns></returns>
        private List<PriceModel> GetListPrices(IEnumerable<ModelExel> cityReport)
        {
            List<PriceModel> prices= new List<PriceModel>();
            foreach (var product in cityReport)
            {
                foreach (var item in product.Prices)
                {
                    prices.Add(item); 
                }
            }
            return prices;
        }
    }
}
