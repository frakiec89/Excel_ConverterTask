namespace Excel_Сonverter.BL
{
    public class ModelExel
    {
        public int Position { get; set; }

        public string ID { get; set; }
        public string NameCity { get; set; }
        public string NameProduct { get; set; }
        public string Unit {  get; set; }
        public List<PriceModel> Prices { get; set; } = new List<PriceModel>();
    }

    public class PriceModel
    {

        public int  Id { get; set; }    
        public double Value { get; set; }
        public string PriceIndicator { get; set; }
        public string Supplier { get; set; }
    }
}