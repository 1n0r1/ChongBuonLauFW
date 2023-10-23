using DocumentFormat.OpenXml;
using Microsoft.VisualBasic.FileIO;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChongBuonLauFW
{
    public partial class StartForm : Form
    {
        private int findFlightKHRRType = 0;
        private BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        private BackgroundWorker backgroundWorker2 = new BackgroundWorker();
        private BackgroundWorker backgroundWorker3 = new BackgroundWorker();
        private BackgroundWorker backgroundWorker4 = new BackgroundWorker();
        private BackgroundWorker backgroundWorker5 = new BackgroundWorker();

        private List<AirportData> airports;
        public StartForm()
        {
            InitializeComponent();

            dataGridView1.CellDoubleClick += DataGridView1_CellDoubleClick;
            dataGridView2.CellDoubleClick += DataGridView2_CellDoubleClick;
            dataGridView3.CellDoubleClick += DataGridView3_CellDoubleClick;
            dataGridView4.CellDoubleClick += DataGridView4_CellDoubleClick;
            dataGridView5.CellDoubleClick += DataGridView5_CellDoubleClick;
            dataGridView6.CellDoubleClick += DataGridView6_CellDoubleClick;

            dataGridView2.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(dataGridView2_RowPrePaint);
            dataGridView6.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(dataGridView6_RowPrePaint);

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;

            backgroundWorker2.WorkerReportsProgress = true;
            backgroundWorker2.DoWork += BackgroundWorker2_DoWork;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorker2_RunWorkerCompleted;

            backgroundWorker3.WorkerReportsProgress = true;
            backgroundWorker3.DoWork += BackgroundWorker3_DoWork;
            backgroundWorker3.RunWorkerCompleted += BackgroundWorker3_RunWorkerCompleted;

            backgroundWorker4.WorkerReportsProgress = true;
            backgroundWorker4.DoWork += BackgroundWorker4_DoWork;
            backgroundWorker4.RunWorkerCompleted += BackgroundWorker4_RunWorkerCompleted;

            backgroundWorker5.WorkerReportsProgress = true;
            backgroundWorker5.DoWork += BackgroundWorker5_DoWork;
            backgroundWorker5.RunWorkerCompleted += BackgroundWorker5_RunWorkerCompleted;

            var collection = MongoUserCollection.GetAirportsCollection();
            var filter = Builders<AirportData>.Filter.Empty;
            airports = collection.Find(filter).ToList();
            List<string> uniqueCountries = airports
                .Select(airport => (airport.CountryName))
                .Distinct()
                .OrderBy(country => country)
                .ToList();
            uniqueCountries.Insert(0, "");
            comboBox1.Items.AddRange(uniqueCountries.ToArray());
            comboBox2.Items.AddRange(uniqueCountries.ToArray());
        }


        private void DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }

        private void DataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView2.Rows.Count)
            {
                DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }

        private void DataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView3.Rows.Count)
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }

        private void DataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView4.Rows.Count)
            {
                DataGridViewRow row = dataGridView4.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }

        private void DataGridView5_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView5.Rows.Count)
            {
                DataGridViewRow row = dataGridView5.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }

        private void DataGridView6_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView6.Rows.Count)
            {
                DataGridViewRow row = dataGridView6.Rows[e.RowIndex];
                string id = row.Cells["Số giấy tờ"].Value?.ToString();
                Form2 form = new Form2(id);
                form.Show();
            }
        }
        private void FindFlight()
        {
            button4.Enabled = false;

            decimal threshold = numericUpDown2.Value;
            bool checkFlight = false;
            string flightNo = textBox3.Text;

            if (flightNo != "") checkFlight = true;

            DateTime rawDay = dateTimePicker4.Value;
            DateTime today = new DateTime(rawDay.Year, rawDay.Month, rawDay.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);

            string filterCountry = comboBox1.Text;

            var filterStage = new BsonDocument("$match", new BsonDocument
            {
                {
                    "FlightList", new BsonDocument
                    {
                        {
                            "$elemMatch", new BsonDocument
                            {
                                {
                                    "Date", new BsonDocument
                                    {
                                        { "$gte", today },
                                        { "$lt", today.AddDays(1) }
                                    }
                                }
                            }
                        }
                    }
                }
            });

            if (checkFlight)
            {
                filterStage = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList", new BsonDocument
                        {
                            {
                                "$elemMatch", new BsonDocument
                                {
                                    {
                                        "Date", new BsonDocument
                                        {
                                            { "$gte", today },
                                            { "$lt", today.AddDays(1) }
                                        }
                                    },
                                    {
                                        "FlightNumber", flightNo
                                    }
                                }
                            }
                        }
                    }
                });
            }

            var filterStageCountry = new BsonDocument
            {
                { "$addFields", new BsonDocument { { "dummy", "hi" } } }
            };

            if (!string.IsNullOrEmpty(filterCountry))
            {
                List<string> codesWithCountryCode = airports
                    .Where(airport => airport.CountryName == filterCountry)
                    .Select(airport => airport.Code)
                    .ToList();

                foreach (string code in codesWithCountryCode)
                {
                    Console.WriteLine(code);
                }

                filterStageCountry = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList", new BsonDocument
                        {
                            {
                                "$elemMatch", new BsonDocument
                                {
                                    {
                                        "Date", new BsonDocument
                                        {
                                            { "$gte", today },
                                            { "$lt", today.AddDays(1) }
                                        }
                                    },
                                    {
                                        "Origin", new BsonDocument
                                        (
                                            "$in", new BsonArray(codesWithCountryCode)
                                        )
                                    }
                                }
                            }
                        }
                    }
                });
            }

            var unwindStage = new BsonDocument("$unwind", new BsonDocument
            {
                {
                    "path", "$FlightList"
                },
                {
                    "preserveNullAndEmptyArrays", false
                }
            });


            var filterStage2 = new BsonDocument("$match", new BsonDocument
            {
                {
                    "FlightList.Date", new BsonDocument
                    {
                        { "$gte", today },
                        { "$lt", today.AddDays(1) }
                    }
                }
            });

            if (checkFlight)
            {
                filterStage2 = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList.Date", new BsonDocument
                        {
                            { "$gte", today },
                            { "$lt", today.AddDays(1) }
                        }
                    },
                    {
                        "FlightList.FlightNumber", flightNo
                    }

                });
            }

            var addFieldsStage1 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "LuggageArray", new BsonDocument
                    {
                        {
                            "$split", new BsonArray {"$FlightList.Luggage", ";"}
                        }
                    }
                }
            });

            var addFieldsStage2 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "LuggageCountArr", new BsonDocument
                    {
                        {
                            "$map", new BsonDocument
                            {
                                {"input", "$LuggageArray"},
                                {"in", new BsonDocument
                                    {
                                        {"$cond", new BsonArray
                                            {
                                                new BsonDocument("$regexMatch", new BsonDocument
                                                {
                                                    {"input", "$$this"},
                                                    {"regex", ",\\d+$"}
                                                }),
                                                new BsonDocument("$toInt", new BsonDocument
                                                    {
                                                        { "$arrayElemAt", new BsonArray
                                                            {
                                                                new BsonDocument
                                                                {
                                                                    { "$split", new BsonArray
                                                                        {
                                                                            "$$this", ","
                                                                        }
                                                                    }
                                                                },
                                                                1
                                                            }
                                                        }
                                                    }
                                                ),
                                                1
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });

            var addFieldsStage3 = new BsonDocument("$addFields", new BsonDocument
            {
                {"LuggageCount", new BsonDocument
                    {
                        {"$subtract", new BsonArray
                            {
                                {new BsonDocument("$sum", "$LuggageCountArr")},
                                1
                            }
                        }
                    }
                }
            });

            var groupStage = new BsonDocument("$group", new BsonDocument
            {
                {"_id", new BsonDocument
                    {
                        {"FlightDate", "$FlightList.Date"},
                        {"FlightNumber", "$FlightList.FlightNumber"},
                        {"Destination", "$FlightList.Destination"},
                        {"Origin", "$FlightList.Origin"},
                        {"Luggage", "$FlightList.Luggage"},
                        {"LuggageCount", "$LuggageCount"}
                    }
                },
                {"Count", new BsonDocument("$sum", 1)},
                {"Data", new BsonDocument("$push", new BsonDocument
                    {
                        {"Name", "$Name"},
                        {"Sex", "$Sex"},
                        {"Nationality", "$Nationality"},
                        {"DOB", "$DOB"},
                        {"IdType", "$IdType"},
                        {"IdNum", "$IdNum"},
                        {"IdProv", "$IdProv"},
                        {"Seat", "$FlightList.Seat"}
                    }
                )}
            });


            var pipeline = new[]
            {
                filterStage,
                filterStageCountry,
                unwindStage,
                filterStage2,
                addFieldsStage1,
                addFieldsStage2,
                addFieldsStage3,
                new BsonDocument("$match", new BsonDocument { { "LuggageCount", new BsonDocument("$gte", threshold) } }),
                new BsonDocument("$sort", new BsonDocument { { "LuggageCount", -1 } }),
                groupStage,
                new BsonDocument("$sort", new BsonDocument { { "_id.LuggageCount", -1 }, { "_id.Luggage", 1 }  })
            };

            if (!backgroundWorker2.IsBusy)
            {
                backgroundWorker2.RunWorkerAsync(pipeline);
            }
        }

        private void FindFlightKHRR()
        {
            button12.Enabled = false;
            button17.Enabled = false;

            DateTime rawDay = dateTimePicker6.Value;
            if (findFlightKHRRType == 1)
            {
                rawDay = dateTimePicker7.Value;
            }
            DateTime today = new DateTime(rawDay.Year, rawDay.Month, rawDay.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);

            DateTime from = DateTime.MinValue;
            DateTime to = DateTime.MaxValue;

            var collectionRR = MongoUserCollection.GetDSRRCollection(findFlightKHRRType);
            var DSRR = collectionRR.Find(Builders<BsonDocument>.Filter.Empty).ToList();
            var filterIdNum = new List<string>();
            foreach (var document in DSRR)
            {
                filterIdNum.Add(document["IdNum"].AsString);
            }

            var filterStage = new BsonDocument("$match", new BsonDocument
            {
                {
                    "IdNum", new BsonDocument
                    {
                        {
                            "$in", new BsonArray(filterIdNum)
                        }
                    }
                },
                {
                    "FlightList", new BsonDocument
                    {
                        {
                            "$elemMatch", new BsonDocument
                            {
                                {
                                    "Date", new BsonDocument
                                    {
                                        { "$gte", today },
                                        { "$lt", today.AddDays(1) }
                                    }
                                }
                            }
                        }
                    }
                }
            });


            var addFieldsStage1 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "Locations", new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input",
                                new BsonDocument("$map",
                                    new BsonDocument
                                        {
                                            { "input",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$FlightList"},
                                                        { "as", "f" },
                                                        { "cond",
                                                            new BsonDocument("$and",
                                                                new BsonArray
                                                                {
                                                                    new BsonDocument("$lt", new BsonArray { "$$f.Date", to}),
                                                                    new BsonDocument("$gte", new BsonArray { "$$f.Date", from})
                                                                }
                                                            )
                                                        }
                                                    }
                                                )
                                            },
                                            { "as", "flight" },
                                            { "in",
                                                new BsonArray
                                                {
                                                    "$$flight.Destination",
                                                    "$$flight.Origin"
                                                }
                                            }
                                        }
                                )
                            },
                            { "initialValue", new BsonArray() },
                            { "in",
                                new BsonDocument("$concatArrays",
                                    new BsonArray
                                        {
                                            "$$this",
                                            "$$value"
                                        }
                                )
                            }
                        }
                    )
                }
            }
            );

            var countRepeatedStage = new BsonDocument("$addFields",
                new BsonDocument("Repeated",
                    new BsonDocument("$arrayToObject",
                        new BsonDocument("$map",
                            new BsonDocument
                            {
                                { "input", new BsonDocument("$setUnion", "$Locations") },
                                { "as", "j" },
                                { "in", new BsonDocument
                                    {
                                        { "k", "$$j" },
                                        { "v",
                                            new BsonDocument("$size",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$Locations" },
                                                        { "cond", new BsonDocument("$eq", new BsonArray
                                                            {
                                                                "$$this",
                                                                "$$j"
                                                            }
                                                        )}
                                                    }
                                                )
                                            )
                                        }
                                    }
                                }
                            }
                        )
                    )
                )
            );

            var addMaxValueStage = new BsonDocument("$addFields",
                new BsonDocument("maxValue",
                    new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input", new BsonDocument("$objectToArray", "$Repeated") },
                            { "initialValue", 0 },
                            { "in",
                                new BsonDocument("$cond",
                                    new BsonArray
                                    {
                                        new BsonDocument("$gte",
                                            new BsonArray
                                            {
                                                "$$this.v",
                                                "$$value"
                                            }
                                        ),
                                        "$$this.v",
                                        "$$value"
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var addFlightDetailsStage = new BsonDocument("$addFields",
                new BsonDocument("FlightDetails",
                    new BsonDocument("$filter",
                        new BsonDocument
                        {
                            { "input", "$FlightList" },
                            { "as", "flight" },
                            { "cond",
                                new BsonDocument("$and",
                                    new BsonArray
                                    {
                                        new BsonDocument("$lt", new BsonArray { "$$flight.Date", today.AddDays(1)}),
                                        new BsonDocument("$gte", new BsonArray { "$$flight.Date", today})
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var pipeline = new[]
            {
                filterStage,
                addFieldsStage1,
                countRepeatedStage,
                addMaxValueStage,
                addFlightDetailsStage,
                new BsonDocument("$sort", new BsonDocument { { "maxValue", -1 }})
            };

            if (!backgroundWorker4.IsBusy)
            {
                backgroundWorker4.RunWorkerAsync(pipeline);
            }
        }

        private void FindFlightKH()
        {
            button9.Enabled = false;

            decimal threshold = numericUpDown3.Value;
            bool checkFlight = false;
            string flightNo = textBox4.Text;

            if (flightNo != "") checkFlight = true;

            DateTime rawDay = dateTimePicker5.Value;
            DateTime today = new DateTime(rawDay.Year, rawDay.Month, rawDay.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);

            DateTime from = DateTime.MinValue;
            DateTime to = DateTime.MaxValue;

            var filterCountry = comboBox2.Text;

            var filterStage = new BsonDocument("$match", new BsonDocument
            {
                {
                    "FlightList", new BsonDocument
                    {
                        {
                            "$elemMatch", new BsonDocument
                            {
                                {
                                    "Date", new BsonDocument
                                    {
                                        { "$gte", today },
                                        { "$lt", today.AddDays(1) }
                                    }
                                }
                            }
                        }
                    }
                }
            });

            if (checkFlight)
            {
                filterStage = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList", new BsonDocument
                        {
                            {
                                "$elemMatch", new BsonDocument
                                {
                                    {
                                        "Date", new BsonDocument
                                        {
                                            { "$gte", today },
                                            { "$lt", today.AddDays(1) }
                                        }
                                    },
                                    {
                                        "FlightNumber", flightNo
                                    }
                                }
                            }
                        }
                    }
                });
            }


            var filterStageCountry = new BsonDocument
            {
                { "$addFields", new BsonDocument { { "dummy", "hi" } } }
            };

            if (!string.IsNullOrEmpty(filterCountry))
            {
                List<string> codesWithCountryCode = airports
                    .Where(airport => airport.CountryName == filterCountry)
                    .Select(airport => airport.Code)
                    .ToList();

                foreach (string code in codesWithCountryCode)
                {
                    Console.WriteLine(code);
                }

                filterStageCountry = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList", new BsonDocument
                        {
                            {
                                "$elemMatch", new BsonDocument
                                {
                                    {
                                        "Date", new BsonDocument
                                        {
                                            { "$gte", today },
                                            { "$lt", today.AddDays(1) }
                                        }
                                    },
                                    {
                                        "Origin", new BsonDocument
                                        (
                                            "$in", new BsonArray(codesWithCountryCode)
                                        )
                                    }
                                }
                            }
                        }
                    }
                });
            }

            var addFieldsStage1 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "Locations", new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input",
                                new BsonDocument("$map",
                                    new BsonDocument
                                        {
                                            { "input",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$FlightList"},
                                                        { "as", "f" },
                                                        { "cond",
                                                            new BsonDocument("$and",
                                                                new BsonArray
                                                                {
                                                                    new BsonDocument("$lt", new BsonArray { "$$f.Date", to}),
                                                                    new BsonDocument("$gte", new BsonArray { "$$f.Date", from})
                                                                }
                                                            )
                                                        }
                                                    }
                                                )
                                            },
                                            { "as", "flight" },
                                            { "in",
                                                new BsonArray
                                                {
                                                    "$$flight.Destination",
                                                    "$$flight.Origin"
                                                }
                                            }
                                        }
                                )
                            },
                            { "initialValue", new BsonArray() },
                            { "in",
                                new BsonDocument("$concatArrays",
                                    new BsonArray
                                        {
                                            "$$this",
                                            "$$value"
                                        }
                                )
                            }
                        }
                    )
                }
            }
            );

            var countRepeatedStage = new BsonDocument("$addFields",
                new BsonDocument("Repeated",
                    new BsonDocument("$arrayToObject",
                        new BsonDocument("$map",
                            new BsonDocument
                            {
                                { "input", new BsonDocument("$setUnion", "$Locations") },
                                { "as", "j" },
                                { "in", new BsonDocument
                                    {
                                        { "k", "$$j" },
                                        { "v",
                                            new BsonDocument("$size",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$Locations" },
                                                        { "cond", new BsonDocument("$eq", new BsonArray
                                                            {
                                                                "$$this",
                                                                "$$j"
                                                            }
                                                        )}
                                                    }
                                                )
                                            )
                                        }
                                    }
                                }
                            }
                        )
                    )
                )
            );

            var addMaxValueStage = new BsonDocument("$addFields",
                new BsonDocument("maxValue",
                    new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input", new BsonDocument("$objectToArray", "$Repeated") },
                            { "initialValue", 0 },
                            { "in",
                                new BsonDocument("$cond",
                                    new BsonArray
                                    {
                                        new BsonDocument("$gte",
                                            new BsonArray
                                            {
                                                "$$this.v",
                                                "$$value"
                                            }
                                        ),
                                        "$$this.v",
                                        "$$value"
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var addFlightDetailsStage = new BsonDocument("$addFields",
                new BsonDocument("FlightDetails",
                    new BsonDocument("$filter",
                        new BsonDocument
                        {
                            { "input", "$FlightList" },
                            { "as", "flight" },
                            { "cond",
                                new BsonDocument("$and",
                                    new BsonArray
                                    {
                                        new BsonDocument("$lt", new BsonArray { "$$flight.Date", today.AddDays(1)}),
                                        new BsonDocument("$gte", new BsonArray { "$$flight.Date", today})
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var pipeline = new[]
            {
                filterStage,
                filterStageCountry,
                addFieldsStage1,
                countRepeatedStage,
                addMaxValueStage,
                addFlightDetailsStage,
                new BsonDocument("$match", new BsonDocument { { "maxValue", new BsonDocument("$gte", threshold) } }),
                new BsonDocument("$sort", new BsonDocument { { "maxValue", -1 }})
            };

            if (!backgroundWorker3.IsBusy)
            {
                backgroundWorker3.RunWorkerAsync(pipeline);
            }
        }

        private void FindAndDisplay()
        {
            button1.Enabled = false;
            DateTime from = DateTime.MinValue;
            DateTime to = DateTime.MaxValue;
            if (checkBox2.Checked)
            {
                DateTime rawDayFrom = dateTimePicker1.Value;
                from = new DateTime(rawDayFrom.Year, rawDayFrom.Month, rawDayFrom.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);
                DateTime rawDayTo = dateTimePicker3.Value;
                to = new DateTime(rawDayTo.Year, rawDayTo.Month, rawDayTo.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7).AddDays(1);
            }

            decimal threshold = numericUpDown1.Value;
            bool hasToday = checkBox1.Checked;
            bool checkRoute = false;
            string node1 = textBox1.Text;
            string node2 = textBox2.Text;
            if (node1 != "" && node2 != "") checkRoute = true;

            var filterStage = new BsonDocument("$match", new BsonDocument { });
            if (hasToday)
            {
                DateTime rawDay = dateTimePicker3.Value;
                DateTime today = new DateTime(rawDay.Year, rawDay.Month, rawDay.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);
                filterStage = new BsonDocument("$match", new BsonDocument
                {
                    {
                        "FlightList", new BsonDocument
                        {
                            {
                                "$elemMatch", new BsonDocument
                                {
                                    {
                                        "Date", new BsonDocument
                                        {
                                            { "$gte", today },
                                            { "$lt", today.AddDays(1) }
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }


            var projectStage = new BsonDocument("$project",
                new BsonDocument
                {
                        { "_id", 0 },
                        { "IdNum", 1 },
                        { "Name", 1 },
                        { "Sex", 1 },
                        { "Nationality", 1 },
                        { "DOB", 1 },
                        { "IdType", 1 },
                        { "IdProv", 1 },
                        { "Locations",
                            new BsonDocument("$reduce",
                                new BsonDocument
                                {
                                    { "input",
                                        new BsonDocument("$map",
                                            new BsonDocument
                                                {
                                                    { "input",
                                                        new BsonDocument("$filter",
                                                            new BsonDocument
                                                            {
                                                                { "input", "$FlightList"},
                                                                { "as", "f" },
                                                                { "cond",
                                                                    new BsonDocument("$and",
                                                                        new BsonArray
                                                                        {
                                                                            new BsonDocument("$lt", new BsonArray { "$$f.Date", to}),
                                                                            new BsonDocument("$gte", new BsonArray { "$$f.Date", from})
                                                                        }
                                                                    )
                                                                }
                                                            }
                                                        )
                                                    },
                                                    { "as", "flight" },
                                                    { "in",
                                                        new BsonArray
                                                        {
                                                            "$$flight.Destination",
                                                            "$$flight.Origin"
                                                        }
                                                    }
                                                }
                                        )
                                    },
                                    { "initialValue", new BsonArray() },
                                    { "in",
                                        new BsonDocument("$concatArrays",
                                            new BsonArray
                                                {
                                                    "$$this",
                                                    "$$value"
                                                }
                                        )
                                    }
                                }
                            )
                        }
                }
            );

            if (checkRoute)
            {
                projectStage = new BsonDocument("$project",
                    new BsonDocument
                        {
                        { "_id", 0 },
                        { "IdNum", 1 },
                        { "Name", 1 },
                        { "Sex", 1 },
                        { "Nationality", 1 },
                        { "DOB", 1 },
                        { "IdType", 1 },
                        { "IdProv", 1 },
                        { "Locations",
                            new BsonDocument("$reduce",
                                new BsonDocument
                                {
                                    { "input",
                                        new BsonDocument("$map",
                                            new BsonDocument
                                                {
                                                    { "input",
                                                        new BsonDocument("$filter",
                                                            new BsonDocument
                                                            {
                                                                { "input", "$FlightList"},
                                                                { "as", "f" },
                                                                { "cond",
                                                                    new BsonDocument("$and",
                                                                        new BsonArray
                                                                        {
                                                                            new BsonDocument("$lt", new BsonArray { "$$f.Date", to}),
                                                                            new BsonDocument("$gte", new BsonArray { "$$f.Date", from}),
                                                                            new BsonDocument("$or",
                                                                                new BsonArray
                                                                                {
                                                                                    new BsonDocument("$and",
                                                                                        new BsonArray
                                                                                        {
                                                                                            new BsonDocument("$eq", new BsonArray { "$$f.Destination", node1}),
                                                                                            new BsonDocument("$eq", new BsonArray { "$$f.Origin", node2})
                                                                                        }
                                                                                    ),
                                                                                    new BsonDocument("$and",
                                                                                        new BsonArray
                                                                                        {
                                                                                            new BsonDocument("$eq", new BsonArray { "$$f.Destination", node2}),
                                                                                            new BsonDocument("$eq", new BsonArray { "$$f.Origin", node1})
                                                                                        }
                                                                                    )
                                                                                }
                                                                            )
                                                                        }
                                                                    )
                                                                }
                                                            }
                                                        )
                                                    },
                                                    { "as", "flight" },
                                                    { "in",
                                                        new BsonArray
                                                        {
                                                            "$$flight.Destination",
                                                            "$$flight.Origin"
                                                        }
                                                    }
                                                }
                                        )
                                    },
                                    { "initialValue", new BsonArray() },
                                    { "in",
                                        new BsonDocument("$concatArrays",
                                            new BsonArray
                                                {
                                                    "$$this",
                                                    "$$value"
                                                }
                                        )
                                    }
                                }
                            )
                        }
                        }
                );
            }


            var countRepeatedStage = new BsonDocument("$addFields",
                new BsonDocument("Repeated",
                    new BsonDocument("$arrayToObject",
                        new BsonDocument("$map",
                            new BsonDocument
                            {
                                { "input", new BsonDocument("$setUnion", "$Locations") },
                                { "as", "j" },
                                { "in", new BsonDocument
                                    {
                                        { "k", "$$j" },
                                        { "v",
                                            new BsonDocument("$size",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$Locations" },
                                                        { "cond", new BsonDocument("$eq", new BsonArray
                                                            {
                                                                "$$this",
                                                                "$$j"
                                                            }
                                                        )}
                                                    }
                                                )
                                            )
                                        }
                                    }
                                }
                            }
                        )
                    )
                )
            );

            var addMaxValueStage = new BsonDocument("$addFields",
                new BsonDocument("maxValue",
                    new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input", new BsonDocument("$objectToArray", "$Repeated") },
                            { "initialValue", 0 },
                            { "in",
                                new BsonDocument("$cond",
                                    new BsonArray
                                    {
                                        new BsonDocument("$gte",
                                            new BsonArray
                                            {
                                                "$$this.v",
                                                "$$value"
                                            }
                                        ),
                                        "$$this.v",
                                        "$$value"
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var pipeline = new[]
            {
                filterStage,
                projectStage,
                countRepeatedStage,
                addMaxValueStage,
                new BsonDocument("$match", new BsonDocument { { "maxValue", new BsonDocument("$gte", threshold) } }),
                new BsonDocument("$sort", new BsonDocument { { "maxValue", -1 }})
            };

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync(pipeline);
            }
        }

        private void FindFlightFrequency()
        {
            button20.Enabled = false;
            decimal threshold = numericUpDown4.Value;

            DateTime rawDay = dateTimePicker8.Value;
            DateTime today = new DateTime(rawDay.Year, rawDay.Month, rawDay.Day, 0, 0, 0, DateTimeKind.Utc).AddHours(-7);

            DateTime from = DateTime.MinValue;
            DateTime to = DateTime.MaxValue;


            DateTime from30daysago = today.AddDays(-30);
            DateTime from90daysago = today.AddDays(-90);

            var filterStage = new BsonDocument("$match", new BsonDocument
            {
                {
                    "FlightList", new BsonDocument
                    {
                        {
                            "$elemMatch", new BsonDocument
                            {
                                {
                                    "Date", new BsonDocument
                                    {
                                        { "$gte", today },
                                        { "$lt", today.AddDays(1) }
                                    }
                                }
                            }
                        }
                    }
                }
            });


            var addFieldsStage1 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "Locations", new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input",
                                new BsonDocument("$map",
                                    new BsonDocument
                                        {
                                            { "input",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$FlightList"},
                                                        { "as", "f" },
                                                        { "cond",
                                                            new BsonDocument("$and",
                                                                new BsonArray
                                                                {
                                                                    new BsonDocument("$lt", new BsonArray { "$$f.Date", to}),
                                                                    new BsonDocument("$gte", new BsonArray { "$$f.Date", from})
                                                                }
                                                            )
                                                        }
                                                    }
                                                )
                                            },
                                            { "as", "flight" },
                                            { "in",
                                                new BsonArray
                                                {
                                                    "$$flight.Destination",
                                                    "$$flight.Origin"
                                                }
                                            }
                                        }
                                )
                            },
                            { "initialValue", new BsonArray() },
                            { "in",
                                new BsonDocument("$concatArrays",
                                    new BsonArray
                                        {
                                            "$$this",
                                            "$$value"
                                        }
                                )
                            }
                        }
                    )
                },
                { "FlightsPast90Days",
                    new BsonDocument("$size",
                        new BsonDocument("$filter",
                            new BsonDocument
                            {
                                { "input", "$FlightList" },
                                { "as", "f" },
                                { "cond",
                                    new BsonDocument("$and",
                                        new BsonArray
                                        {
                                            new BsonDocument("$lt", new BsonArray { "$$f.Date", to }),
                                            new BsonDocument("$gte", new BsonArray { "$$f.Date", from90daysago })
                                        }
                                    )
                                }
                            }
                        )
                    )
                },
                { "FlightsPast30Days",
                    new BsonDocument("$size",
                        new BsonDocument("$filter",
                            new BsonDocument
                            {
                                { "input", "$FlightList" },
                                { "as", "f" },
                                { "cond",
                                    new BsonDocument("$and",
                                        new BsonArray
                                        {
                                            new BsonDocument("$lt", new BsonArray { "$$f.Date", to }),
                                            new BsonDocument("$gte", new BsonArray { "$$f.Date", from30daysago })
                                        }
                                    )
                                }
                            }
                        )
                    )
                }
            }
            );

            var addFieldsStage2 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "FlightsPast90DaysAverage", new BsonDocument("$divide",
                        new BsonArray{"$FlightsPast90Days", 3 }
                     )
                }
            }
            );

            var addFieldsStage3 = new BsonDocument("$addFields", new BsonDocument
            {
                {
                    "FlightsFrequencyDifference", new BsonDocument("$subtract",
                        new BsonArray
                        {
                            new BsonDocument("$multiply", new BsonArray {"$FlightsPast30Days", 3}),
                            "$FlightsPast90Days"
                        }
                    )
                }
            }
            );



            var countRepeatedStage = new BsonDocument("$addFields",
                new BsonDocument("Repeated",
                    new BsonDocument("$arrayToObject",
                        new BsonDocument("$map",
                            new BsonDocument
                            {
                                { "input", new BsonDocument("$setUnion", "$Locations") },
                                { "as", "j" },
                                { "in", new BsonDocument
                                    {
                                        { "k", "$$j" },
                                        { "v",
                                            new BsonDocument("$size",
                                                new BsonDocument("$filter",
                                                    new BsonDocument
                                                    {
                                                        { "input", "$Locations" },
                                                        { "cond", new BsonDocument("$eq", new BsonArray
                                                            {
                                                                "$$this",
                                                                "$$j"
                                                            }
                                                        )}
                                                    }
                                                )
                                            )
                                        }
                                    }
                                }
                            }
                        )
                    )
                )
            );

            var addMaxValueStage = new BsonDocument("$addFields",
                new BsonDocument("maxValue",
                    new BsonDocument("$reduce",
                        new BsonDocument
                        {
                            { "input", new BsonDocument("$objectToArray", "$Repeated") },
                            { "initialValue", 0 },
                            { "in",
                                new BsonDocument("$cond",
                                    new BsonArray
                                    {
                                        new BsonDocument("$gte",
                                            new BsonArray
                                            {
                                                "$$this.v",
                                                "$$value"
                                            }
                                        ),
                                        "$$this.v",
                                        "$$value"
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var addFlightDetailsStage = new BsonDocument("$addFields",
                new BsonDocument("FlightDetails",
                    new BsonDocument("$filter",
                        new BsonDocument
                        {
                            { "input", "$FlightList" },
                            { "as", "flight" },
                            { "cond",
                                new BsonDocument("$and",
                                    new BsonArray
                                    {
                                        new BsonDocument("$lt", new BsonArray { "$$flight.Date", today.AddDays(1)}),
                                        new BsonDocument("$gte", new BsonArray { "$$flight.Date", today})
                                    }
                                )
                            }
                        }
                    )
                )
            );

            var pipeline = new[]
            {
                filterStage,
                addFieldsStage1,
                new BsonDocument("$match", new BsonDocument { { "FlightsPast30Days", new BsonDocument("$gte", threshold) } }),
                addFieldsStage2,
                addFieldsStage3,
                countRepeatedStage,
                addMaxValueStage,
                addFlightDetailsStage,
                // new BsonDocument("$match", new BsonDocument { { "FlightsFrequencyDifference", new BsonDocument("$gte", 1) } }),
                new BsonDocument("$sort", new BsonDocument { { "FlightsPast90Days", -1 }})
            };

            if (!backgroundWorker5.IsBusy)
            {
                backgroundWorker5.RunWorkerAsync(pipeline);
            }
        }


        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var pipeline = e.Argument as BsonDocument[];

            if (worker == null || pipeline == null)
            {
                // Handle null or invalid arguments
                return;
            }

            var collection = MongoUserCollection.GetMongoUserCollection();
            var result = collection.Aggregate<BsonDocument>(pipeline).ToList();

            // Create a new DataTable to store the results
            var dataTable = new DataTable();
            dataTable.Columns.Add("Số giấy tờ");
            dataTable.Columns.Add("Họ tên");
            dataTable.Columns.Add("Giới tính");
            dataTable.Columns.Add("Quốc tịch");
            dataTable.Columns.Add("Ngày sinh");
            dataTable.Columns.Add("Loại giấy tờ");
            dataTable.Columns.Add("Nơi cấp");
            dataTable.Columns.Add("Tuyến đường");
            dataTable.Columns.Add("Số lần nhiều nhất");
            dataTable.Columns["Số lần nhiều nhất"].DataType = typeof(Int32);

            foreach (var doc in result)
            {
                var row = dataTable.NewRow();
                row["Số giấy tờ"] = doc.GetValue("IdNum", string.Empty);
                row["Họ tên"] = doc.GetValue("Name", string.Empty);
                row["Giới tính"] = doc.GetValue("Sex", string.Empty);
                row["Quốc tịch"] = doc.GetValue("Nationality", string.Empty);
                row["Ngày sinh"] = doc.GetValue("DOB", string.Empty);
                row["Loại giấy tờ"] = doc.GetValue("IdType", string.Empty);
                row["Nơi cấp"] = doc.GetValue("IdProv", string.Empty);
                var repeatedField = doc.GetValue("Repeated", new BsonDocument()).AsBsonDocument;
                var repeatedString = string.Empty;

                foreach (var element in repeatedField)
                {
                    var key = element.Name;
                    var value = element.Value.ToString();
                    repeatedString += $"{key}: {value}, ";
                }
                row["Tuyến đường"] = repeatedString;
                row["Số lần nhiều nhất"] = doc.GetValue("maxValue", 0).ToInt32();
                dataTable.Rows.Add(row);

            }

            e.Result = dataTable;
        }
        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button1.Enabled = true;
            if (e.Error != null)
            {
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!e.Cancelled)
            {
                var dataTable = e.Result as DataTable;
                dataGridView1.DataSource = dataTable;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    if (i < dataGridView1.Rows.Count)
                    {
                        dataGridView1.Rows[i].HeaderCell.Value = "" + (i + 1);
                    }
                }
            }
        }

        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var pipeline = e.Argument as BsonDocument[];

            if (worker == null || pipeline == null)
            {
                return;
            }

            var collection = MongoUserCollection.GetMongoUserCollection();
            var result = collection.Aggregate<BsonDocument>(pipeline).ToList();

            // Create a new DataTable to store the results
            var dataTable = new DataTable();
            dataTable.Columns.Add("Ngày bay");
            dataTable.Columns.Add("Chuyến bay");
            dataTable.Columns.Add("Số ghế");
            dataTable.Columns.Add("Họ tên");
            dataTable.Columns.Add("Giới tính");
            dataTable.Columns.Add("Quốc tịch");
            dataTable.Columns.Add("Ngày sinh");
            dataTable.Columns.Add("Loại giấy tờ");
            dataTable.Columns.Add("Số giấy tờ");
            dataTable.Columns.Add("Nơi cấp");
            dataTable.Columns.Add("Nơi đi");
            dataTable.Columns.Add("Nơi đến");
            dataTable.Columns.Add("Hành lý");
            dataTable.Columns.Add("Số hành lý");
            dataTable.Columns["Số hành lý"].DataType = typeof(Int32);
            dataTable.Columns.Add("Số khách");
            dataTable.Columns.Add("Data");
            dataTable.Columns.Add("Color");

            string currentcolor = "white";

            int totalLuggage = 0;
            foreach (var doc in result)
            {
                TimeZoneInfo vstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                DateTime flightdate = TimeZoneInfo.ConvertTimeFromUtc(doc["_id"]["FlightDate"].ToUniversalTime(), vstZone);

                totalLuggage += doc["_id"]["LuggageCount"].AsInt32;
                foreach (var u in doc["Data"].AsBsonArray)
                {
                    string seat = u["Seat"].ToString();
                    string name = u["Name"].ToString();
                    string sex = u["Sex"].ToString();
                    string nat = u["Nationality"].ToString();
                    string dob = u["DOB"].ToString();
                    string idtype = u["IdType"].ToString();
                    string idnum = u["IdNum"].ToString();
                    string idprov = u["IdProv"].ToString();

                    var row = dataTable.NewRow();
                    row["Ngày bay"] = flightdate.ToString("dd/MM/yyyy");
                    row["Chuyến bay"] = doc["_id"]["FlightNumber"];
                    row["Số ghế"] = seat;
                    row["Họ tên"] = name;
                    row["Giới tính"] = sex;
                    row["Quốc tịch"] = nat;
                    row["Ngày sinh"] = dob;
                    row["Loại giấy tờ"] = idtype;
                    row["Số giấy tờ"] = idnum;
                    row["Nơi cấp"] = idprov;
                    row["Nơi đi"] = doc["_id"]["Origin"];
                    row["Nơi đến"] = doc["_id"]["Destination"];
                    row["Hành lý"] = doc["_id"]["Luggage"];

                    row["Số hành lý"] = doc["_id"]["LuggageCount"];
                    row["Số khách"] = doc["Count"];
                    row["Data"] = doc["Data"];
                    row["Color"] = currentcolor;

                    dataTable.Rows.Add(row);
                }
                if (currentcolor == "white") currentcolor = "grey";
                else currentcolor = "white";

            }

            var customResult = new Tuple<DataTable, int>(dataTable, totalLuggage);
            e.Result = customResult;
        }
        private void BackgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button4.Enabled = true;
            if (e.Error != null)
            {
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!e.Cancelled)
            {
                var customResult = e.Result as Tuple<DataTable, int>;

                var dataTable = customResult.Item1;
                dataGridView2.DataSource = dataTable;
                dataGridView2.Columns["Data"].Visible = false;
                dataGridView2.Columns["Color"].Visible = false;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    if (i < dataGridView2.Rows.Count)
                    {
                        dataGridView2.Rows[i].HeaderCell.Value = "" + (i + 1);
                    }
                }

                var totalLuggage = customResult.Item2;
                label12.Text = "Tổng số hành lý: " + totalLuggage;
                label13.Text = "Trung bình hành lý mỗi khách: " + (float) totalLuggage/ dataTable.Rows.Count;
                label14.Text = "Số khách: " + dataTable.Rows.Count;

            }
        }

        private void BackgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var pipeline = e.Argument as BsonDocument[];

            if (worker == null || pipeline == null)
            {
                return;
            }

            var collection = MongoUserCollection.GetMongoUserCollection();
            var result = collection.Aggregate<BsonDocument>(pipeline).ToList();

            // Create a new DataTable to store the results
            var dataTable = new DataTable();
            dataTable.Columns.Add("Ngày bay");
            dataTable.Columns.Add("Chuyến bay");
            dataTable.Columns.Add("Số ghế");
            dataTable.Columns.Add("Họ tên");
            dataTable.Columns.Add("Giới tính");
            dataTable.Columns.Add("Quốc tịch");
            dataTable.Columns.Add("Ngày sinh");
            dataTable.Columns.Add("Loại giấy tờ");
            dataTable.Columns.Add("Số giấy tờ");
            dataTable.Columns.Add("Nơi cấp");
            dataTable.Columns.Add("Nơi đi");
            dataTable.Columns.Add("Nơi đến");
            dataTable.Columns.Add("Hành lý");

            dataTable.Columns.Add("Tuyến đường");
            dataTable.Columns.Add("Số lần nhiều nhất");


            foreach (var doc in result)
            {
                TimeZoneInfo vstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                DateTime flightdate = TimeZoneInfo.ConvertTimeFromUtc(doc["FlightDetails"][0]["Date"].ToUniversalTime(), vstZone);

                var row = dataTable.NewRow();

                row["Số giấy tờ"] = doc.GetValue("IdNum", string.Empty);
                row["Họ tên"] = doc.GetValue("Name", string.Empty);
                row["Giới tính"] = doc.GetValue("Sex", string.Empty);
                row["Quốc tịch"] = doc.GetValue("Nationality", string.Empty);
                row["Ngày sinh"] = doc.GetValue("DOB", string.Empty);
                row["Loại giấy tờ"] = doc.GetValue("IdType", string.Empty);
                row["Nơi cấp"] = doc.GetValue("IdProv", string.Empty);
                var repeatedField = doc.GetValue("Repeated", new BsonDocument()).AsBsonDocument;
                var repeatedString = string.Empty;

                foreach (var element in repeatedField)
                {
                    var key = element.Name;
                    var value = element.Value.ToString();
                    repeatedString += $"{key}: {value}, ";
                }
                row["Tuyến đường"] = repeatedString;
                row["Số lần nhiều nhất"] = doc.GetValue("maxValue", 0).ToInt32();

                row["Nơi đi"] = doc["FlightDetails"][0]["Origin"];
                row["Nơi đến"] = doc["FlightDetails"][0]["Destination"];
                row["Hành lý"] = doc["FlightDetails"][0]["Luggage"];
                row["Ngày bay"] = flightdate.ToString("dd/MM/yyyy");
                row["Chuyến bay"] = doc["FlightDetails"][0]["FlightNumber"];
                row["Số ghế"] = doc["FlightDetails"][0]["Seat"];

                dataTable.Rows.Add(row);
            }


            e.Result = dataTable;
        }
        private void BackgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button9.Enabled = true;
            if (e.Error != null)
            {
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!e.Cancelled)
            {
                var dataTable = e.Result as DataTable;
                dataGridView3.DataSource = dataTable;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    if (i < dataGridView3.Rows.Count)
                    {
                        dataGridView3.Rows[i].HeaderCell.Value = "" + (i + 1);
                    }
                }
            }
        }

        private void BackgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var pipeline = e.Argument as BsonDocument[];

            if (worker == null || pipeline == null)
            {
                return;
            }

            var collection = MongoUserCollection.GetMongoUserCollection();
            var result = collection.Aggregate<BsonDocument>(pipeline).ToList();

            // Create a new DataTable to store the results
            var dataTable = new DataTable();
            dataTable.Columns.Add("Ngày bay");
            dataTable.Columns.Add("Chuyến bay");
            dataTable.Columns.Add("Số ghế");
            dataTable.Columns.Add("Họ tên");
            dataTable.Columns.Add("Giới tính");
            dataTable.Columns.Add("Quốc tịch");
            dataTable.Columns.Add("Ngày sinh");
            dataTable.Columns.Add("Loại giấy tờ");
            dataTable.Columns.Add("Số giấy tờ");
            dataTable.Columns.Add("Nơi cấp");
            dataTable.Columns.Add("Nơi đi");
            dataTable.Columns.Add("Nơi đến");
            dataTable.Columns.Add("Hành lý");

            dataTable.Columns.Add("Tuyến đường");
            dataTable.Columns.Add("Số lần nhiều nhất");
            dataTable.Columns.Add("Ghi chú");


            foreach (var doc in result)
            {
                TimeZoneInfo vstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                DateTime flightdate = TimeZoneInfo.ConvertTimeFromUtc(doc["FlightDetails"][0]["Date"].ToUniversalTime(), vstZone);

                var row = dataTable.NewRow();

                row["Số giấy tờ"] = doc.GetValue("IdNum", string.Empty);
                row["Ghi chú"] = doc.GetValue("Note", string.Empty);
                if ((string)row["Ghi chú"] == "BsonNull")
                {
                    row["Ghi chú"] = "";
                }
                row["Họ tên"] = doc.GetValue("Name", string.Empty);
                row["Giới tính"] = doc.GetValue("Sex", string.Empty);
                row["Quốc tịch"] = doc.GetValue("Nationality", string.Empty);
                row["Ngày sinh"] = doc.GetValue("DOB", string.Empty);
                row["Loại giấy tờ"] = doc.GetValue("IdType", string.Empty);
                row["Nơi cấp"] = doc.GetValue("IdProv", string.Empty);
                var repeatedField = doc.GetValue("Repeated", new BsonDocument()).AsBsonDocument;
                var repeatedString = string.Empty;

                foreach (var element in repeatedField)
                {
                    var key = element.Name;
                    var value = element.Value.ToString();
                    repeatedString += $"{key}: {value}, ";
                }
                row["Tuyến đường"] = repeatedString;
                row["Số lần nhiều nhất"] = doc.GetValue("maxValue", 0).ToInt32();

                row["Nơi đi"] = doc["FlightDetails"][0]["Origin"];
                row["Nơi đến"] = doc["FlightDetails"][0]["Destination"];
                row["Hành lý"] = doc["FlightDetails"][0]["Luggage"];
                row["Ngày bay"] = flightdate.ToString("dd/MM/yyyy");
                row["Chuyến bay"] = doc["FlightDetails"][0]["FlightNumber"];
                row["Số ghế"] = doc["FlightDetails"][0]["Seat"];

                dataTable.Rows.Add(row);
            }

            e.Result = dataTable;
        }

        private void BackgroundWorker4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button12.Enabled = true;
            button17.Enabled = true;
            if (e.Error != null)
            {
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!e.Cancelled)
            {
                var dataTable = e.Result as DataTable;
                if (findFlightKHRRType == 0)
                {
                    dataGridView4.DataSource = dataTable;
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        if (i < dataGridView4.Rows.Count)
                        {
                            dataGridView4.Rows[i].HeaderCell.Value = "" + (i + 1);
                        }
                    }
                }
                else
                {
                    dataGridView5.DataSource = dataTable;
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        if (i < dataGridView5.Rows.Count)
                        {
                            dataGridView5.Rows[i].HeaderCell.Value = "" + (i + 1);
                        }
                    }
                }
            }
        }


        private void BackgroundWorker5_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            var pipeline = e.Argument as BsonDocument[];

            if (worker == null || pipeline == null)
            {
                return;
            }

            var collection = MongoUserCollection.GetMongoUserCollection();
            var result = collection.Aggregate<BsonDocument>(pipeline).ToList();

            // Create a new DataTable to store the results
            var dataTable = new DataTable();
            dataTable.Columns.Add("Ngày bay");
            dataTable.Columns.Add("Chuyến bay");
            dataTable.Columns.Add("Số ghế");
            dataTable.Columns.Add("Họ tên");
            dataTable.Columns.Add("Giới tính");
            dataTable.Columns.Add("Quốc tịch");
            dataTable.Columns.Add("Ngày sinh");
            dataTable.Columns.Add("Loại giấy tờ");
            dataTable.Columns.Add("Số giấy tờ");
            dataTable.Columns.Add("Nơi cấp");
            dataTable.Columns.Add("Nơi đi");
            dataTable.Columns.Add("Nơi đến");
            dataTable.Columns.Add("Hành lý");

            dataTable.Columns.Add("Tuyến đường");
            dataTable.Columns.Add("Số lần nhiều nhất");

            dataTable.Columns.Add("Số chuyến trong vòng 90 ngày");
            dataTable.Columns.Add("Số chuyến trung bình mỗi 30 ngày");
            dataTable.Columns.Add("Số chuyến trong vòng 30 ngày");

            dataTable.Columns.Add("Color");

            foreach (var doc in result)
            {
                TimeZoneInfo vstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                DateTime flightdate = TimeZoneInfo.ConvertTimeFromUtc(doc["FlightDetails"][0]["Date"].ToUniversalTime(), vstZone);

                var row = dataTable.NewRow();

                row["Số giấy tờ"] = doc.GetValue("IdNum", string.Empty);
                row["Họ tên"] = doc.GetValue("Name", string.Empty);
                row["Giới tính"] = doc.GetValue("Sex", string.Empty);
                row["Quốc tịch"] = doc.GetValue("Nationality", string.Empty);
                row["Ngày sinh"] = doc.GetValue("DOB", string.Empty);
                row["Loại giấy tờ"] = doc.GetValue("IdType", string.Empty);
                row["Nơi cấp"] = doc.GetValue("IdProv", string.Empty);
                var repeatedField = doc.GetValue("Repeated", new BsonDocument()).AsBsonDocument;
                var repeatedString = string.Empty;

                foreach (var element in repeatedField)
                {
                    var key = element.Name;
                    var value = element.Value.ToString();
                    repeatedString += $"{key}: {value}, ";
                }
                row["Tuyến đường"] = repeatedString;
                row["Số lần nhiều nhất"] = doc.GetValue("maxValue", 0).ToInt32();

                row["Nơi đi"] = doc["FlightDetails"][0]["Origin"];
                row["Nơi đến"] = doc["FlightDetails"][0]["Destination"];
                row["Hành lý"] = doc["FlightDetails"][0]["Luggage"];
                row["Ngày bay"] = flightdate.ToString("dd/MM/yyyy");
                row["Chuyến bay"] = doc["FlightDetails"][0]["FlightNumber"];
                row["Số ghế"] = doc["FlightDetails"][0]["Seat"];

                row["Số chuyến trong vòng 90 ngày"] = doc["FlightsPast90Days"];
                row["Số chuyến trung bình mỗi 30 ngày"] = (float)doc["FlightsPast90Days"].AsInt32 / 3;
                row["Số chuyến trong vòng 30 ngày"] = doc["FlightsPast30Days"];
                if (doc["FlightsPast30Days"].AsInt32 > (float)doc["FlightsPast90Days"].AsInt32 / 3)
                {
                    row["Color"] = "yellow";
                }
                else
                {
                    row["Color"] = "white";
                }

                dataTable.Rows.Add(row);
            }

            e.Result = dataTable;
        }

        private void BackgroundWorker5_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button20.Enabled = true;
            if (e.Error != null)
            {
                MessageBox.Show($"Error: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!e.Cancelled)
            {
                var dataTable = e.Result as DataTable;
                dataGridView6.DataSource = dataTable;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    if (i < dataGridView6.Rows.Count)
                    {
                        dataGridView6.Rows[i].HeaderCell.Value = "" + (i + 1);
                    }
                }
                dataGridView6.Columns["Color"].Visible = false;
                dataGridView6.Columns["Tuyến đường"].Visible = false;
                dataGridView6.Columns["Số lần nhiều nhất"].Visible = false;
            }
        }


        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];

            if (row.Cells["Color"].Value != null && row.Cells["Color"].Value.ToString().Contains("white"))
            {
                row.DefaultCellStyle.BackColor = Color.White;
            }
            else
            {
                row.DefaultCellStyle.BackColor = Color.LightGray;
            }
        }

        private void dataGridView6_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow row = dataGridView6.Rows[e.RowIndex];

            if (row.Cells["Color"].Value != null && row.Cells["Color"].Value.ToString().Contains("yellow"))
            {
                row.DefaultCellStyle.BackColor = Color.Yellow;
            }
            else
            {
                row.DefaultCellStyle.BackColor = Color.White;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1();
            form.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FindAndDisplay();
        }

        private void StartForm_Shown(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FindFlight();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView2);

        }

        private void button7_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportForm(dataGridView2);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            FindFlightKH();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportFormHK(dataGridView3);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            findFlightKHRRType = 0;
            FindFlightKHRR();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView4);

        }

        private void button10_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportFormHK(dataGridView4);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Form4 form = new Form4();
            form.Show();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView5);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportFormHK(dataGridView5);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            findFlightKHRRType = 1;
            FindFlightKHRR();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Form5 form = new Form5();
            form.Show();

        }

        private void button20_Click(object sender, EventArgs e)
        {
            FindFlightFrequency();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportToExcel(dataGridView6);
        }

        private void button18_Click(object sender, EventArgs e)
        {
            ExcelExporter.ExportFormHK(dataGridView6);
        }
    }
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

