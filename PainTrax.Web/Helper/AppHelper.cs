using System.Data;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using MS.Services;
using PainTrax.Services;
using System.IO;
//using iTextSharp.text.pdf;
//using iTextSharp.text;

namespace PainTrax.Web.Helper
{
    public class AppHelper
    {
        public static string DataTableToJsonObj(System.Data.DataTable dt)
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Merge(dt);
            StringBuilder JsonString = new StringBuilder();
            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                JsonString.Append("[");
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    JsonString.Append("{");
                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {
                        if (j < ds.Tables[0].Columns.Count - 1)
                        {
                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + ds.Tables[0].Rows[i][j].ToString() + "\",");
                        }
                        else if (j == ds.Tables[0].Columns.Count - 1)
                        {
                            JsonString.Append("\"" + ds.Tables[0].Columns[j].ColumnName.ToString() + "\":" + "\"" + ds.Tables[0].Rows[i][j].ToString() + "\"");
                        }
                    }
                    if (i == ds.Tables[0].Rows.Count - 1)
                    {
                        JsonString.Append("}");
                    }
                    else
                    {
                        JsonString.Append("},");
                    }
                }
                JsonString.Append("]");
                return JsonString.ToString();
            }
            else
            {
                return null;
            }
        }

        public string GetJson(string query)
        {
            ParentService _parentservices = new ParentService();
            System.Data.DataTable dt = _parentservices.GetData(query);
            return DataTableToJsonObj(dt);
        }

        public byte[] GetExcel(DataTable dt)
        {
            // Create a MemoryStream to store the Excel file
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Create a spreadsheet document in the MemoryStream
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    // Add a WorkbookPart to the document
                    var workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    // Add a Sheets to the Workbook
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Appointments" };
                    sheets.Append(sheet);

                    // Get the sheetData of the Worksheet
                    var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    // Add some data to the Excel file
                    var row = new Row();
                    row.AppendChild(new Cell() { CellValue = new CellValue("Date"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Time"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Phone"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Location"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Provider"), DataType = CellValues.String });
                    row.AppendChild(new Cell() { CellValue = new CellValue("Status"), DataType = CellValues.String });
                    sheetData.AppendChild(row);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string date = Convert.ToDateTime(dt.Rows[i]["AppointmentDate"].ToString()).ToString("MM/dd/yyyy");
                        string time = Convert.ToDateTime(dt.Rows[i]["AppointmentDate"].ToString() + " " + dt.Rows[i]["AppointmentStart"].ToString()).ToString("HH:mm tt");
                        string name = dt.Rows[i]["fname"].ToString() + " " + dt.Rows[i]["lname"].ToString() + "\n" + dt.Rows[i]["AppointmentNote"].ToString();
                        string phone = dt.Rows[i]["home_ph"].ToString() == "" ? dt.Rows[i]["mobile"].ToString() : dt.Rows[i]["home_ph"].ToString();
                        string location = dt.Rows[i]["Location"].ToString();
                        string provider = dt.Rows[i]["ProviderName"].ToString();
                        string status = dt.Rows[i]["Status"].ToString();

                        var datarow = new Row();
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(date), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(time), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(name), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(phone), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(location), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(provider), DataType = CellValues.String });
                        datarow.AppendChild(new Cell() { CellValue = new CellValue(status), DataType = CellValues.String });

                        sheetData.AppendChild(datarow);

                    }

                    // Save the changes to the spreadsheet document
                    workbookPart.Workbook.Save();
                }
                return memoryStream.ToArray();
            }

        }

        public byte[] GetPdf(DataTable dt)
        {
            // Create a MemoryStream to store the Excel file
            MemoryStream pdfOutput = new MemoryStream();
            //Document document = new Document(PageSize.A4, 25, 25, 30, 30);

            //PdfWriter writer = PdfWriter.GetInstance(document, pdfOutput);
            //document.AddAuthor("Paintrax");
            //document.AddCreator("Paintrax");
            //document.AddKeywords("");
            //document.AddSubject("Autogenrated  PDF document");
            //document.AddTitle("PDF Report");
            //document.Open();

            //iTextSharp.text.Paragraph p = new iTextSharp.text.Paragraph("Appointments\n\n", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10.0f, 1));

            ////p.SetAlignment("Center");
            //document.Add(p);
            //float[] columnWidth = { 80.0f, 70.0f, 200.0f, 90.0f, 70.0f, 100.0f, 100.0f };
            ////float[] columnWidth = { 80.0f, 70.0f, 200.0f,  70.0f, 100.0f, 100.0f };

            //PdfPTable table = new PdfPTable(columnWidth);
            //table.AddCell(new Paragraph("Date", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Time", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Name", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Phone", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Location", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Provider", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));
            //table.AddCell(new Paragraph("Status", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9.0f, 1)));


            //for (int i = 0; i < dt.Rows.Count; i++)
            //{


            //    table.AddCell(new Paragraph(Convert.ToDateTime(dt.Rows[i]["AppointmentDate"].ToString()).ToString("MM/dd/yyyy"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    string time = Convert.ToDateTime(dt.Rows[i]["AppointmentDate"].ToString() + " " + dt.Rows[i]["AppointmentStart"].ToString()).ToString("HH:mm tt");
            //    //time += " To ";
            //    //time += Convert.ToDateTime(dt.Rows[i]["AppointmentDate"].ToString() + " " + dt.Rows[i]["AppointmentEnd"].ToString()).ToString("HH:mm tt");
            //    table.AddCell(new Paragraph(time, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    table.AddCell(new Paragraph(dt.Rows[i]["FName"].ToString() + " " + dt.Rows[i]["LName"].ToString() + "\n" + dt.Rows[i]["AppointmentNote"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    table.AddCell(new Paragraph(dt.Rows[i]["home_ph"].ToString() == "" ? dt.Rows[i]["mobile"].ToString() : dt.Rows[i]["home_ph"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    table.AddCell(new Paragraph(dt.Rows[i]["Location"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    table.AddCell(new Paragraph(dt.Rows[i]["ProviderName"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));
            //    table.AddCell(new Paragraph(dt.Rows[i]["Status"].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8.0f, 0)));



            //}
            //document.Add(table);
            //document.Close();
            //writer.Close();
            ////fs.Close();
            return pdfOutput.ToArray();

        }

    }
}
