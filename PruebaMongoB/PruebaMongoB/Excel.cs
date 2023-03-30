using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using MongoDB.Bson;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PruebaMongoB
{
    public class Excel
    {
        public static List<BsonDocument> ReadExcel(string filePath)
        {
            List<BsonDocument> documents = new List<BsonDocument>();

            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            IWorkbook workbook = new XSSFWorkbook(fs);

            ISheet sheet = workbook.GetSheetAt(0);

            int rowsCount = sheet.PhysicalNumberOfRows;

            for (int i = 1; i < rowsCount; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                    continue;
                int cellsCount = row.Cells.Count;
                BsonDocument document = new BsonDocument();
                for (int j = 0; j < cellsCount; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                        continue;
                    string cellValue = cell.ToString();
                    string cellHeader = sheet.GetRow(0).GetCell(j).ToString();
                    document.Add(cellHeader, cellValue);
                }
                documents.Add(document);
            }

            return documents;
        }
    }
}
