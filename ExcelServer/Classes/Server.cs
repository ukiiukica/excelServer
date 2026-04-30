using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;

namespace ExcelServer.Classes
{
    public class Server
    {
        private HttpListener listener;
        public Server()
        {
            listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5050/");
        }
        public void Pokreni()
        {
            listener.Start();
            while (true)
            {
                HttpListenerContext context = listener.GetContext();
                HttpListenerRequest request = context.Request;
                HttpListenerResponse response = context.Response;

                Uri uri = request.Url;
                string csvFajl = uri.AbsolutePath;
                if (csvFajl == "/favicon.ico")
                {
                    response.StatusCode = 204;
                    response.Close();
                    continue;
                }
                else if (!csvFajl.EndsWith(".csv"))
                {
                    using (var writer = new StreamWriter(response.OutputStream))
                    {
                        writer.Write("URL ne sadrzi CSV fajl");
                    }
                    response.Close();
                }
                else
                {
                    string basePath = AppDomain.CurrentDomain.BaseDirectory;
                    string path = Path.Combine(basePath, csvFajl.Remove(0,1));
                    Console.WriteLine(path);
                    try
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook();
                        var sheet = workbook.CreateSheet("Sheet1");

                        string[] lines = File.ReadAllLines(path);

                        for (int i = 0; i < lines.Length; i++)
                        {
                            var row = sheet.CreateRow(i);
                            var values = lines[i].Split(',');
                            for (int j = 0; j < values.Length; j++)
                            {
                                row.CreateCell(j).SetCellValue(values[j]);
                            }
                        }

                        using (var excelBytes = new MemoryStream())
                        {
                            workbook.Write(excelBytes);
                            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            response.AddHeader("Content-Disposition", "attachment; filename=output.xlsx");


                            response.OutputStream.Write(excelBytes.ToArray(), 0, excelBytes.ToArray().Length);
                        }
                    }
                    catch(Exception ex)
                    {
                        response.StatusCode = 404;
                        using (var writer = new StreamWriter(response.OutputStream))
                        {
                            writer.Write($"Ne postoji fajl {csvFajl} na serveru");
                        }
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        response.Close();
                    }
                }
            }
        }
    }
}
