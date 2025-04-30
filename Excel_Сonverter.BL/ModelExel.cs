namespace Excel_Сonverter.BL
{
    public class ModelExel
    {
        public int Position { get; set; }
        public string ID { get; set; }
        public string NameCity { get; set; }
        public string CodeCity { get; set; }
        
        public string NameProduct { get; set; }
        public string Unit {  get; set; }
        public List<PriceModel> Prices { get; set; } = new List<PriceModel>();
    }

    public class PriceModel
    {
        private string priceIndicator;

        public int  Id { get; set; }    
        public double Value { get; set; }
        public string PriceIndicator 
        {   get { return priceIndicator; }

            set
            {
                if (value.Trim().ToLower() == "жц" || value.Trim().ToLower() == "т")     // Сделайем ограничение
                     priceIndicator = value; 
                else
                    priceIndicator =string.Empty;
            } 
        }
        public string Supplier { get; set; }
    }
}