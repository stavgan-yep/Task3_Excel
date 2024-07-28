using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;


namespace Task3_Excel
{
    public class ExcelData
    {
        public List<Product> Products { get; set; }
        public List<Client> Clients { get; set; }
        public List<Order> Orders { get; set; }

        public ExcelData(string filePath)
        {
            Products = new List<Product>();
            Clients = new List<Client>();
            Orders = new List<Order>();

            using (var workbook = new XLWorkbook(filePath))
            {
                LoadProducts(workbook);
                LoadClients(workbook);
                LoadOrders(workbook);
            }
        }

        private void LoadProducts(IXLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet("Товары");
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                Products.Add(new Product
                {
                    ProductId = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Unit = row.Cell(3).GetValue<string>(),
                    Price = row.Cell(4).GetValue<decimal>()
                });
            }
        }

        private void LoadClients(IXLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet("Клиенты");
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                Clients.Add(new Client
                {
                    ClientId = row.Cell(1).GetValue<int>(),
                    OrganizationName = row.Cell(2).GetValue<string>(),
                    Address = row.Cell(3).GetValue<string>(),
                    ContactPerson = row.Cell(4).GetValue<string>()
                });
            }
        }

        private void LoadOrders(IXLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet("Заявки");
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                Orders.Add(new Order
                {
                    OrderId = row.Cell(1).GetValue<int>(),
                    ProductId = row.Cell(2).GetValue<int>(),
                    ClientId = row.Cell(3).GetValue<int>(),
                    OrderNumber = row.Cell(4).GetValue<string>(),
                    Quantity = row.Cell(5).GetValue<int>(),
                    OrderDate = row.Cell(6).GetValue<DateTime>()
                });
            }
        }

        public void SaveChanges(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("Клиенты");
                foreach (var client in Clients)
                {

                    var row = worksheet.RowsUsed().Skip(1).FirstOrDefault(r => r.Cell(1).GetValue<int>() == client.ClientId);
                    if (row != null)
                    {
                        row.Cell(4).Value = client.ContactPerson;
                    }
                }
                workbook.Save();
            }
        }
    }
}
