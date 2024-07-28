using System;
using System.IO;
using System.Linq;

namespace Task3_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = GetValidFilePath();

            var excelData = new ExcelData(filePath);

            while (true)
            {
                Console.WriteLine("Выберите команду (ввести только номер команды):");
                Console.WriteLine("1. Информация о клиентах по наименованию товара");
                Console.WriteLine("2. Изменение контактного лица клиента");
                Console.WriteLine("3. Определение золотого клиента");
                Console.WriteLine("4. Выход");
                string command = Console.ReadLine();

                if (command == "1")
                {
                    Console.WriteLine("Введите наименование товара:");
                    string productName = Console.ReadLine();
                    GetCustomerInfo(excelData, productName);
                }
                else if (command == "2")
                {
                    Console.WriteLine("Введите название организации:");
                    string organizationName = Console.ReadLine();
                    Console.WriteLine("Введите ФИО нового контактного лица:");
                    string newContactPerson = Console.ReadLine();
                    UpdateContactPerson(excelData, filePath, organizationName, newContactPerson);
                }
                else if (command == "3")
                {
                    Console.WriteLine("Введите год:");
                    int year = int.Parse(Console.ReadLine());
                    Console.WriteLine("Введите месяц:");
                    int month = int.Parse(Console.ReadLine());
                    GetGoldenCustomer(excelData, year, month);
                }
                else if (command == "4")
                {
                    break;
                }
                else
                {
                    Console.WriteLine("Неверная команда, попробуйте снова.");
                }
            }
        }

        static void GetCustomerInfo(ExcelData excelData, string productName)
        {
            var product = excelData.Products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));
            if (product != null)
            {
                var orders = excelData.Orders.Where(o => o.ProductId == product.ProductId);
                foreach (var order in orders)
                {
                    var client = excelData.Clients.First(c => c.ClientId == order.ClientId);
                    Console.WriteLine($"Клиент: {client.OrganizationName}, Количество: {order.Quantity}, Цена: {product.Price}, Дата заказа: {order.OrderDate.ToShortDateString()}");
                }
                if (orders.Count() == 0)
                {
                    Console.WriteLine($"Клиенты с товаром {productName} не найдены.");
                }
            }
            else
            {
                Console.WriteLine($"Товар с наименованием {productName} не найден.");
            }
        }

        static void UpdateContactPerson(ExcelData excelData, string filePath, string organizationName, string newContactPerson)
        {
            var client = excelData.Clients.FirstOrDefault(c => c.OrganizationName.Equals(organizationName, StringComparison.OrdinalIgnoreCase));
            if (client != null)
            {
                client.ContactPerson = newContactPerson;
                excelData.SaveChanges(filePath);
                Console.WriteLine($"Контактное лицо для организации {organizationName} изменено на {newContactPerson}.");
            }
            else
            {
                Console.WriteLine($"Организация {organizationName} не найдена.");
            }
        }

        static void GetGoldenCustomer(ExcelData excelData, int year, int month)
        {
            var orders = excelData.Orders.Where(o => o.OrderDate.Year == year && o.OrderDate.Month == month);
            var customerOrders = orders.GroupBy(o => o.ClientId)
                                       .Select(g => new { ClientId = g.Key, OrderCount = g.Count() })
                                       .OrderByDescending(g => g.OrderCount)
                                       .FirstOrDefault();

            if (customerOrders != null)
            {
                var client = excelData.Clients.First(c => c.ClientId == customerOrders.ClientId);
                Console.WriteLine($"Золотой клиент за {year}-{month}: {client.OrganizationName} с {customerOrders.OrderCount} заказами.");
            }
            else
            {
                Console.WriteLine($"Нет данных о заказах за {year}-{month}.");
            }
        }

        static string GetValidFilePath()
        {
            while (true)
            {
                Console.WriteLine("Введите путь до файла Excel:");
                string filePath = Console.ReadLine();

                if (File.Exists(filePath))
                {
                    Console.WriteLine("Файл найден. Продолжаем работу...");
                    return filePath;
                }
                else
                {
                    Console.WriteLine("Файл не найден. Попробуйте снова.");
                }
            }
        }
    }
}





