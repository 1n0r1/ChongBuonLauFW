using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Core.Authentication;
using System;
using System.Windows.Forms;

namespace ChongBuonLauFW
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            LoginForm loginForm = new LoginForm();
            DialogResult result = loginForm.ShowDialog();

            if (result == DialogResult.OK)
            {
                Application.Run(new StartForm());
            }
            else
            {
                MessageBox.Show("Login failed. Exiting application.");
            }
        }

    }
    public static class DatabaseMongoCollection
    {
        private static MongoClient client = null;
        public static void Login(string username, string password)
        {
            client = new MongoClient($"mongodb+srv://{username}:{password}@cluster0.76k4h.mongodb.net/?retryWrites=true&w=majority");
            try
            {
                var databaseNames = client.ListDatabaseNames().ToList();
            }
            catch (MongoAuthenticationException ex)
            {
                Console.WriteLine($"Authentication failed: {ex.Message}");
                MessageBox.Show("Không đúng tài khoản hoặc mật khẩu");
                throw new ApplicationException("Authentication failed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                MessageBox.Show("Không kết nối được. Có thể vì địa chỉ IP không được cho phép.");
                throw new ApplicationException("Connection failed");
            }
        }

        private static IMongoDatabase GetDatabase()
        {
            // MongoClient client = new MongoClient("mongodb+srv://test:nzhrwUuhk37OXhZz@cluster1.pcviqfw.mongodb.net/?retryWrites=true&w=majority");
            return client.GetDatabase("database");
        }
        public static IMongoCollection<Person> GetMongoUserCollection()
        {
            IMongoCollection<Person> collection = GetDatabase().GetCollection<Person>("User");
            return collection;

        }

        public static IMongoCollection<BsonDocument> GetDSRRCollection(int type)
        {
            IMongoCollection<BsonDocument> collection = GetDatabase().GetCollection<BsonDocument>("DSRR");
            if (type == 1)
            {
                collection = GetDatabase().GetCollection<BsonDocument>("DSRROp");
            }
            return collection;

        }
        public static IMongoCollection<AirportData> GetAirportsCollection()
        {
            IMongoCollection<AirportData> collection = GetDatabase().GetCollection<AirportData>("Airports");
            return collection;

        }
    }
}
