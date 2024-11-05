namespace DomainToExcel.Domain
{
    public class Domain : Base
    {
        public int Id { get; set; }
        public DirectionReference Direction { get; set; }
        public AddressReference Address { get; set; }
        public List<Hours> Hours { get; set; }

    }
}
