using System;

using MongoDB.Bson;
using System.Collections.Generic;

namespace ChongBuonLauFW
{
    public class Person
    {
        public ObjectId _id { get; set; }
        public string IdNum { get; set; }
        public string Name { get; set; }
        public string Sex { get; set; }
        public string Nationality { get; set; }
        public string DOB { get; set; }
        public string IdType { get; set; }
        public string IdProv { get; set; }
        public string Note { get; set; }
        public List<Flight> FlightList { get; set; }
    }
    public class Flight
    {
        public string Origin { get; set; }
        public string Destination { get; set; }
        public string Luggage { get; set; }
        public DateTime Date { get; set; }
        public string FlightNumber { get; set; }
        public string Seat { get; set; }
    }
    public class AirportData
    {
        public ObjectId _id { get; set; }
        public string Continent { get; set; }
        public string CountryCode { get; set; }
        public string Region { get; set; }
        public string Code { get; set; }
        public string CountryName { get; set; }
    }


}
