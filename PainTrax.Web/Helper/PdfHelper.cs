using DocumentFormat.OpenXml.Vml;
using iTextSharp.text;
using iTextSharp.text.pdf;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Element;
//using iText.Layout.Properties;
//using iText.IO.Font.Constants;
//using iText.Kernel.Font;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Routing.Constraints;
using PainTrax.Services;
using System.Data;
using System.Text;


namespace PainTrax.Web.Helper
{
    public class PdfHelper
    {
        public byte[] GeneratePdf(string html)
        {
            byte[] pdfBytes;

            using (var memoryStream = new MemoryStream())
            {
                var document = new Document();
                var writer = PdfWriter.GetInstance(document, memoryStream);

                document.Open();
                using (var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
                {
                    //  XMLWorkerHelper.GetInstance().ParseXHtml(writer, document, htmlStream, null, Encoding.UTF8);
                }
                document.Close();

                pdfBytes = memoryStream.ToArray();
            }

            return pdfBytes;
        }
        public byte[] Stamping(string SourceFile, string ColumnName, string ID, Dictionary<string, string> controls, string cmpid = "0", string path = "")
        {
            WebHostBuilderContext webHost = new WebHostBuilderContext();
            ParentService _parentService = new ParentService();
            PdfReader pdfReader = new PdfReader(SourceFile);
            AcroFields readPdfFields = pdfReader.AcroFields;
            String tabname = readPdfFields.GetField("txtTable");
            if (tabname == null || tabname == "")
                tabname = "View_Pdf";


            DataTable dt = _parentService.GetData("select * from " + tabname + " where " + ColumnName + "=" + ID);

            MemoryStream pdfOutput = new MemoryStream();
            PdfStamper pdfStamper = new PdfStamper(pdfReader, pdfOutput);
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            pdfStamper.FormFlattening = false;
            AcroFields ae = pdfReader.AcroFields;
            foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
            {

                string textvalue = pdfFormFields.GetField(de.Key.ToString());
                string[] textpair = textvalue.Split('|');
                if (textpair.Length > 1 && !de.Key.StartsWith("#"))
                {
                    try
                    {
                        if (dt.Rows[0][textpair[1]] is DateTime)
                            pdfFormFields.SetField(textpair[0], DateTime.Parse(dt.Rows[0][textpair[1]].ToString()).ToString("MM/dd/yyyy"));
                        else
                            pdfFormFields.SetField(textpair[0], dt.Rows[0][textpair[1]].ToString());

                        if (textpair.Length > 2)
                        {
                            if (dt.Rows[0][textpair[1]] == null || dt.Rows[0][textpair[1]].ToString().Trim() == string.Empty)
                                pdfFormFields.SetField(textpair[0], textpair[2]);
                            else
                                pdfFormFields.SetField(textpair[0], "");
                        }

                    }
                    catch (Exception ex)
                    {
                        pdfFormFields.SetField(textpair[0], "");

                    }
                }
                else
                {

                    if (!de.Key.ToLower().StartsWith("txtfix"))
                    {
                        if (controls != null && controls.Count != 0)
                        {
                            if (de.Key.StartsWith("@"))
                            {
                                try
                                {
                                    pdfFormFields.SetField(de.Key.ToString(), controls[de.Key.Substring(1)]);
                                }
                                catch (Exception ex)
                                {
                                    pdfFormFields.SetField(textpair[0], "");
                                }
                            }
                            else
                                pdfFormFields.SetField(textpair[0], "");
                        }
                        else
                        {
                            pdfFormFields.SetField(textpair[0], "");
                        }

                        if (de.Key.StartsWith("#"))
                        {
                            try
                            {
                                string[] headpair = de.Key.Substring(1).Split('|');
                                if (headpair.Length == 1)
                                {
                                    try
                                    {
                                        if (dt.Rows[0][de.Key.Substring(1)] is DateTime)
                                            pdfFormFields.SetField(de.Key.ToString(), DateTime.Parse(dt.Rows[0][de.Key.Substring(1)].ToString()).ToString("MM/dd/yyyy"));
                                        else
                                            pdfFormFields.SetField(de.Key.ToString(), dt.Rows[0][de.Key.Substring(1)].ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        pdfFormFields.SetField(de.Key.Substring(1), "");
                                    }
                                }
                                else if (headpair.Length == 4)
                                {
                                    try
                                    {
                                        DataTable dtCode = _parentService.GetData("select * from " + headpair[0] + " where " + headpair[2] + "='" + controls[headpair[1]] + "' and cmp_id=" + cmpid);
                                        if (dtCode.Rows.Count > 0)
                                        {
                                            pdfFormFields.SetField(de.Key.ToString(), dtCode.Rows[0][headpair[3]].ToString());
                                        }
                                        else
                                        {
                                            pdfFormFields.SetField(de.Key.ToString(), "");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        pdfFormFields.SetField(de.Key.Substring(1), "");
                                    }

                                }
                                else
                                {
                                    pdfFormFields.SetField(textpair[0], "");
                                }
                            }
                            catch (Exception ex)
                            {
                                pdfFormFields.SetField(textpair[0], "");
                            }
                        }
                        else
                            pdfFormFields.SetField(textpair[0], "");
                    }
                    else
                        pdfFormFields.SetField(textpair[0], "");

                }
                if (de.Key.ToLower().StartsWith("imgsign"))
                {

                    try
                    {
                        if (path != "")
                        {
                            //  path = System.IO.Path.Combine(path, "Sign");
                            string[] files = System.IO.Directory.GetFiles(path, ID + ".jp*g", System.IO.SearchOption.TopDirectoryOnly);
                            if (files.Length > 0)
                            {

                                Stream inputImageStream = new FileStream(files[0], FileMode.Open, FileAccess.Read, FileShare.Read);
                                float[] fieldPosition = null;
                                fieldPosition = ae.GetFieldPositions(de.Key.ToString());
                                var pdfContentByte = pdfStamper.GetOverContent((int)fieldPosition[0]);
                                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputImageStream);
                                image.ScaleToFit(fieldPosition[3] - fieldPosition[1], fieldPosition[4] - fieldPosition[2] + 10);
                                image.SetAbsolutePosition(fieldPosition[1], fieldPosition[2]);
                                pdfContentByte.AddImage(image);
                            }
                        }
                    }

                    catch (Exception ex) { }
                    finally { pdfFormFields.SetFieldProperty(de.Key.ToString(), "flags", PdfFormField.FLAGS_NOVIEW, null); }
                }

            }
            pdfStamper.Close();
            pdfReader.Close();
            return pdfOutput.ToArray();

        }


        public void AcknoStamping(string SourceFile, Dictionary<string, string> controls, string path = "", string fileName = "", string signPath = "")
        {
            WebHostBuilderContext webHost = new WebHostBuilderContext();

            PdfReader pdfReader = new PdfReader(SourceFile);
            AcroFields readPdfFields = pdfReader.AcroFields;


            MemoryStream pdfOutput = new MemoryStream();
            PdfStamper pdfStamper = new PdfStamper(pdfReader, pdfOutput);
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            pdfStamper.FormFlattening = false;
            AcroFields ae = pdfReader.AcroFields;
            foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
            {
                if (de.Key.ToLower().StartsWith("imgsign"))
                {

                    try
                    {
                        if (path != "")
                        {

                            if (!string.IsNullOrEmpty(signPath))
                            {

                                Stream inputImageStream = new FileStream(signPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                                float[] fieldPosition = null;
                                fieldPosition = ae.GetFieldPositions(de.Key.ToString());
                                var pdfContentByte = pdfStamper.GetOverContent((int)fieldPosition[0]);
                                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputImageStream);
                                image.ScaleToFit((fieldPosition[3] - fieldPosition[1]) + 50, (fieldPosition[4] - fieldPosition[2]) - 10);
                                image.SetAbsolutePosition(fieldPosition[1] + 50, fieldPosition[2] + 10);
                                pdfContentByte.AddImage(image);
                            }
                        }
                    }

                    catch (Exception ex) { }
                    finally { pdfFormFields.SetFieldProperty(de.Key.ToString(), "flags", PdfFormField.FLAGS_NOVIEW, null); }
                }
                else
                {
                    pdfFormFields.SetField(de.Key.ToString(), controls[de.Key.Substring(1)]);
                }


            }
            pdfStamper.Close();
            pdfReader.Close();
            byte[] pdfBytes = pdfOutput.ToArray();

            // ✅ Save to specific location
            if (!string.IsNullOrEmpty(path))
            {
                // Ensure directory exists
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string fullPath = System.IO.Path.Combine(path, fileName);

                File.WriteAllBytes(fullPath, pdfBytes);
            }



        }


        /*
        public byte[] Stamping(string SourceFile, string ColumnName, string ID, Dictionary<string, string> controls, string cmpid = "0", string path = "")
        {
            WebHostBuilderContext webHost = new WebHostBuilderContext();
            ParentService _parentService = new ParentService();
            PdfReader pdfReader = new PdfReader(SourceFile);
            AcroFields readPdfFields = pdfReader.AcroFields;
            String tabname = readPdfFields.GetField("txtTable");
            //filename = readPdfFields.GetField("txtFile");
            if (tabname == null || tabname == "")
                tabname = "ViewPdf";
            // if (filename == null || filename == "")
            //  filename = "FileName";


            DataTable dt = _parentService.GetData("select * from " + tabname + " where " + ColumnName + "=" + ID);
            string fileprefix = "1";
            //try
            //{
            //    fileprefix = dt.Rows[0][filename].ToString();
            //}
            //catch { }
            //try
            //{
            //    iTextSharp.text.pdf.AcroFields.Item txtExcel = pdfReader.AcroFields.Fields["txtExcel"];
            //    PdfDictionary merged = txtExcel.GetValue(0);
            //    string attrib = "";
            //    foreach (PdfName key in merged.Keys)
            //    {
            //        attrib += key + " " + merged.GetDirectObject(key);
            //    }
            //}
            //catch { }

            MemoryStream pdfOutput = new MemoryStream();
            PdfStamper pdfStamper = new PdfStamper(pdfReader, pdfOutput);
            AcroFields pdfFormFields = pdfStamper.AcroFields;
            pdfStamper.FormFlattening = false;
            AcroFields ae = pdfReader.AcroFields;
            foreach (KeyValuePair<string, iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
            {

                string textvalue = pdfFormFields.GetField(de.Key.ToString());
                string[] textpair = textvalue.Split('|');
                if (textpair.Length > 1)
                {
                    try
                    {
                        if (dt.Rows[0][textpair[1]] is DateTime)
                            pdfFormFields.SetField(textpair[0], DateTime.Parse(dt.Rows[0][textpair[1]].ToString()).ToString("MM/dd/yyyy"));
                        else
                            pdfFormFields.SetField(textpair[0], dt.Rows[0][textpair[1]].ToString());

                        if (textpair.Length > 2)
                        {
                            if (dt.Rows[0][textpair[1]] == null || dt.Rows[0][textpair[1]].ToString().Trim() == string.Empty)
                                pdfFormFields.SetField(textpair[0], textpair[2]);
                            else
                                pdfFormFields.SetField(textpair[0], "");
                        }

                    }
                    catch (Exception ex)
                    {
                        pdfFormFields.SetField(textpair[0], "");

                    }
                }
                else
                {
                    if (!de.Key.ToLower().StartsWith("txtfix"))
                    {
                        if (controls != null && controls.Count != 0)
                        {
                            if (de.Key.StartsWith("@"))
                            {
                                try
                                {
                                    pdfFormFields.SetField(de.Key.ToString(), controls[de.Key.Substring(1)]);
                                    //   TextBox txt = (TextBox)ctrl.FindControl(de.Key.Substring(1));
                                    //  pdfFormFields.SetField(de.Key.ToString(), txt.Text);

                                }
                                catch (Exception ex)
                                {
                                    pdfFormFields.SetField(textpair[0], "");
                                }
                            }
                            else
                                pdfFormFields.SetField(textpair[0], "");
                        }
                        else
                        {
                            pdfFormFields.SetField(textpair[0], "");
                        }


                        if (de.Key.StartsWith("#"))
                        {
                            try
                            {
                                string[] headpair = de.Key.Substring(1).Split('|');
                                if (headpair.Length == 1)
                                {
                                    try
                                    {
                                        if (dt.Rows[0][de.Key.Substring(1)] is DateTime)
                                            pdfFormFields.SetField(de.Key.ToString(), DateTime.Parse(dt.Rows[0][de.Key.Substring(1)].ToString()).ToString("MM/dd/yyyy"));
                                        else
                                            pdfFormFields.SetField(de.Key.ToString(), dt.Rows[0][de.Key.Substring(1)].ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        pdfFormFields.SetField(de.Key.Substring(1), "");
                                    }
                                }
                                else if (headpair.Length == 4)
                                {

                                    // TextBox txt = (TextBox)ctrl.FindControl(headpair[1]);

                                    //                                HttpContext.Current.Response.Write(txt.Text);
                                    //                              HttpContext.Current.Response.Write(headpair[2]);
                                    //                            HttpContext.Current.Response.Write(headpair[3]);
                                    //                          HttpContext.Current.Response.Write("select * from " + headpair[0] + " where " + headpair[2] + "='" + txt.Text + "'");
                                    //DataTable dtCode = GetData("select * from " + headpair[0] + " where " + headpair[2] + "='" + txt.Text + "'");
                                    DataTable dtCode = _parentService.GetData("select * from " + headpair[0] + " where " + headpair[2] + "='" + controls[de.Key.Substring(1)] + "' and cmpid=" + cmpid);
                                    if (dtCode.Rows.Count > 0)
                                    {
                                        pdfFormFields.SetField(de.Key.ToString(), dtCode.Rows[0][headpair[3]].ToString());
                                        //HttpContext.Current.Response.Write(dtCode.Rows [0][ headpair[3]]);
                                    }
                                    else
                                    {
                                        pdfFormFields.SetField(textpair[0], "");
                                    }
                                }
                                else
                                {
                                    pdfFormFields.SetField(textpair[0], "");
                                }
                            }
                            catch (Exception ex)
                            {
                                pdfFormFields.SetField(textpair[0], "");
                            }
                            //try
                            //{
                            //    pdfFormFields.SetField(de.Key.ToString(), dt.Rows[0][de.Key.Substring(1)].ToString());
                            //}
                            //catch (Exception ex)
                            //{
                            //    pdfFormFields.SetField(textpair[0], "");
                            //}
                        }
                    }
                    else
                        pdfFormFields.SetField(textpair[0], "");

                }
                if (de.Key.ToLower().StartsWith("imgsign"))
                {

                    try
                    {
                        // string webRootPath = webHost.HostingEnvironment.WebRootPath;
                        // string path = "";
                        if (path != "")
                        {
                            path = System.IO.Path.Combine(path, "Sign");
                            // DataTable dts = GetData("select * from tblPatientIESign where " + ColumnName + "=" + ID);
                            //string[] files = System.IO.Directory.GetFiles(path, ID + "_*.jp*g", System.IO.SearchOption.TopDirectoryOnly);
                            string[] files = System.IO.Directory.GetFiles(path, "*_*.jp*g", System.IO.SearchOption.TopDirectoryOnly);
                            if (files.Length > 0)
                            {

                                Stream inputImageStream = new FileStream(files[0], FileMode.Open, FileAccess.Read, FileShare.Read);
                                float[] fieldPosition = null;
                                fieldPosition = ae.GetFieldPositions(de.Key.ToString());

                                // AcroFields.FieldPosition f = ae.GetFieldPositions(de.Key.ToString())[0];
                                //AcroFields.FieldPosition[] fieldPositions = ae.GetFieldPositions(de.Key.ToString());
                                //  var pdfContentByte = pdfStamper.GetOverContent(f.page);
                                var pdfContentByte = pdfStamper.GetOverContent((int)fieldPosition[0]);
                                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputImageStream);
                                //image.ScaleToFit(f.position.Width, f.position.Height + 10);

                                //image.SetAbsolutePosition(f.position.Left, f.position.Bottom);
                                image.ScaleToFit(fieldPosition[3] - fieldPosition[1], fieldPosition[4] - fieldPosition[2] + 10);

                                image.SetAbsolutePosition(fieldPosition[1], fieldPosition[2]);
                                pdfContentByte.AddImage(image);
                            }
                        }
                        //    pdfFormFields.SetFieldProperty(de.Key.ToString(), "flags", PdfFormField.FLAGS_NOVIEW, null);
                    }

                    catch (Exception ex) { }
                    finally { pdfFormFields.SetFieldProperty(de.Key.ToString(), "flags", PdfFormField.FLAGS_NOVIEW, null); }
                }

            }
            //try
            //{
            //    AcroFields.FieldPosition f = ae.GetFieldPositions("imgsign")[0];
            //    var pdfContentByte = pdfStamper.GetOverContent(f.page);
            //    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputImageStream);
            //    image.ScaleToFit(f.position.Width, f.position.Height + 10);

            //    image.SetAbsolutePosition(f.position.Left, f.position.Bottom);
            //    pdfContentByte.AddImage(image);

            //    pdfFormFields.SetFieldProperty("imgsign", "flags", PdfFormField.FLAGS_NOVIEW, null);
            //}
            //catch { }
            pdfStamper.Close();
            pdfReader.Close();
            return pdfOutput.ToArray();
            try
            {
                //var response = HttpContext.Current.Response;
                //response.AddHeader("Content-Disposition", "attachment; filename=\"" + fileprefix + "_" + Path.GetFileName(SourceFile) + "\"");
                //response.ContentType = "application/pdf";
                //response.BinaryWrite(pdfOutput.ToArray());
                //response.End();
            }
            catch (Exception ex)
            {
                // Logger.Error(ex);
            }

        }*/
    }
}
