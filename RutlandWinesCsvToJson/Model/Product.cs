using System;

namespace RutlandWinesCsvToJson.Model
{

        public class Product
        {
            public Guid Id { get; set; }

            public string Title { get; set; }

            public string LongDescription { get; set; }

            public string SellingStory { get; set; }

            public string FoodSuggestion { get; set; }

            public string Region { get; set; }

            public string Country { get; set; }

            public double RestaurantPrice { get; set; }

            public int Size { get; set; }

            public string ImageUrl { get; set; }

            public int Vintage { get; set; }

            public double ABV { get; set; }

            public double IndividualBottleSalePrice { get; set; }

            public double CaseSalePrice { get; set; }

            public string WineType { get; set; }

            public bool MilkUsed { get; set; }

            public bool EggUsed { get; set; }

            public bool Sulphites { get; set; }

            public bool Vegatarain { get; set; }

            public bool Vegan { get; set; }

            public bool Organic { get; set; }

            public bool Biodynamic { get; set; }
    }
}

