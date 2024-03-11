using MongoDB.Bson;
using MongoDB.Driver;
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
            Application.Run(new StartForm());
        }

    }
    public static class DatabaseMongoCollection
    {
        private static IMongoDatabase GetDatabase()
        {
            MongoClient client = new MongoClient("mongodb+srv://test:nzhrwUuhk37OXhZz@cluster1.pcviqfw.mongodb.net/?retryWrites=true&w=majority");
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
