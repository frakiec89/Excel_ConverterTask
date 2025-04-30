

namespace Excel_Сonverter.BL
{
    internal class ReportService
    {
        public List<ModelReport> GetModelReports (List<ModelExel> models)
        {
            List<string> cities = models.Select(x=>x.CodeCity).Distinct().ToList();
            List<ModelReport> modelReports = new List<ModelReport>();

            int x = 0;  
            foreach (var codeCity in cities)
            {

                var cityReport = models.Where(x => x.CodeCity == codeCity);

                List<PriceModel> prices = GetListPrices(cityReport);

                x++;
                var m = new ModelReport();
                m.NameCity = cityReport.FirstOrDefault().NameCity;
                m.CodeCity = codeCity;
                m.Position = x; 
                m.CountProduct = prices.Where(x=>x.PriceIndicator.ToLower().Contains("жц")).Count();
                m.CountT = prices.Where(x => x.PriceIndicator.ToLower().Contains("т")).Count();
                m.CountEmpty = prices.Where(x => x.Value==0).Count();
                m.fillPercentage = ((m.CountChZ + m.CountT) / 5) *100;// todo поменять 5  ;
                modelReports.Add(m);
            }
            return modelReports;
        }

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
