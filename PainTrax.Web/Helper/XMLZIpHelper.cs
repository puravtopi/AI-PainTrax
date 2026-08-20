using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace PainTrax.Web.Helper
{
    public class XMLZipHelper
    {
        public byte[] GenerateZip(DataTable dt, string xmlTemplateFile)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string content = File.ReadAllText(xmlTemplateFile);

                        // Replace placeholders
                        foreach (DataColumn col in row.Table.Columns)
                        {
                            string placeholder = $"`{col.ColumnName}`";
                            string value = row[col]?.ToString() ?? "";
                            content = content.Replace(placeholder, value);
                        }

                        string fileName =
                            row["lname"].ToString()+"_"+ row["lname"].ToString() + ".xml";

                        var entry = zip.CreateEntry(fileName);

                        using (var stream = entry.Open())
                        using (var writer = new StreamWriter(stream, Encoding.UTF8))
                        {
                            writer.Write(content);
                        }
                    }
                }

                return ms.ToArray();
            }
        }
    }
}