
namespace Excel_Сonverter.BL
{
    public class ModelReport
    {
        public int Position { get; set; }

        public string NameCity { get; set; }
        public string CodeCity { get; set; }

        public int CountProduct { get; set; }

        /// <summary>
        /// Кол-во позиций с отметкой «ЖЦ»  что бы это не значило
        /// </summary>
        public int  CountChZ { get; set; }
        public int  CountT { get; set; }
        public int  CountEmpty { get; set; }

        public double fillPercentage { get; set; }



    }
}
