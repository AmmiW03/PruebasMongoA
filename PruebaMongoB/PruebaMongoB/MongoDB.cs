using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PruebaMongoB
{
    public class MongoDB
    {
        public static void Insertar(BsonDocument document)
        {
            try
            {
                MongoClient oClient = new MongoClient("mongodb://localhost:27017");

                var oDatabase = oClient.GetDatabase("controlCursos");

                var coleccion = oDatabase.GetCollection<BsonDocument>("datos");
                coleccion.InsertOne(document);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
