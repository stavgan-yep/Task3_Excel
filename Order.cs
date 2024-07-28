using System;

namespace Task3_Excel
{
    public class Order
    {
        public int OrderId { get; set; }
        public int ProductId { get; set; }
        public int ClientId { get; set; }
        public string OrderNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }
}
