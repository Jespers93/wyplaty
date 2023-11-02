using System.ComponentModel.DataAnnotations;

namespace wyplaty.Entities
{
    public class Fracht
    {
        public int Id { get; set; }
        [Required]
        public string OrderNumber { get; set; }
        public DateTime LoadDate { get; set; }
        public DateTime UnloadDate { get; set; }
        public string Car { get; set; }
        public string CityLoad { get; set; }
        public string CityUnload { get; set; }
        public double Price { get; set; }
        virtual public Driver? Driver { get; set; }
       
}
}
