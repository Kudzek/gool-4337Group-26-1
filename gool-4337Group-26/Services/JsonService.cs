using Newtonsoft.Json;
using gool_4337Group_26.Models;
using System.Collections.Generic;
using System.IO;

namespace gool_4337Group_26.Services
{
    public class JsonService
    {
        public List<Client> Import(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Файл не найден: {path}");

            var json = File.ReadAllText(path);
            var clients = JsonConvert.DeserializeObject<List<Client>>(json);

            if (clients == null)
                return new List<Client>();

            // Валидация данных
            foreach (var client in clients)
            {
                if (client.FullName != null)
                    client.FullName = client.FullName.Trim();
                if (client.Email != null)
                    client.Email = client.Email.Trim();
            }

            return clients;
        }
    }
}