using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;

namespace ExcelServer.Classes
{
    public class Server
    {
        private HttpListener listener;
        private static Queue<HttpListenerContext> queue = new Queue<HttpListenerContext>();
        private static object queuelock = new object();
        private Dictionary<string, KesInstanca> kes = new Dictionary<string, KesInstanca>();
        private static readonly object kesLock = new object();
        private static readonly object logLock = new object();
        private static readonly object genLock = new object();
        private int aktivneNiti = 0;
        private static readonly object obradaLock = new object();
        private bool shutdown = false;
        private int maxQueueCount = 5;
        public Server()
        {
            listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5050/");
        }
        public void Pokreni()
        {
            listener.Start();

            for (int i = 0; i < 5; i++)
                ThreadPool.QueueUserWorkItem(Obrada);

            new Thread(GracefulShutdown).Start();

            while (!shutdown)
            {
                try
                {
                    HttpListenerContext context = listener.GetContext();
                    lock (queuelock)
                    {
                        if (queue.Count +  aktivneNiti <= maxQueueCount)
                        {
                            queue.Enqueue(context);
                            Monitor.Pulse(queuelock);
                        }
                        else
                        {
                            ServerPreopterecen(context);
                        }
                    }
                }
                catch(HttpListenerException e)
                {
                    if (shutdown)
                        break;
                    throw;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    break;
                }
            }
        }
        private void Obrada(object state)
        {
            while (true)
            {
                HttpListenerContext context;

                lock (queuelock)
                {
                    while (queue.Count == 0)
                    {
                        if (shutdown)
                            return;
                        Monitor.Wait(queuelock);
                    }
                    context = queue.Dequeue();
                    lock (obradaLock)
                    {
                        aktivneNiti++;
                    }
                }

                PronadjiCSV(context);
                lock (obradaLock)
                {
                    aktivneNiti--;

                    if (shutdown && aktivneNiti == 0 && queue.Count == 0)
                    {
                        Monitor.PulseAll(obradaLock);
                    }
                }
            }
        }
        public void PronadjiCSV(HttpListenerContext context)
        {
            var start = System.Diagnostics.Stopwatch.StartNew();

            HttpListenerRequest request = context.Request;
            HttpListenerResponse response = context.Response;

            Uri uri = request.Url;
            string csvFajl = uri.AbsolutePath;

            if (csvFajl == "/favicon.ico")
            {
                response.StatusCode = 204;
                response.Close();
            }
            else if (!csvFajl.EndsWith(".csv"))
            {
                response.StatusCode = 400;
                using (var writer = new StreamWriter(response.OutputStream))
                {
                    writer.Write("URL ne sadrzi CSV fajl");
                }
                response.Close();
            }
            else
            {
                try
                {
                    csvFajl = csvFajl.Remove(0, 1);
                    string imeFajla = csvFajl.Substring(0, csvFajl.Length - 4);

                    byte[] buffer = null;

                    string poruka="";

                    KesInstanca instanca = null;
                    lock (kesLock)
                    {
                        kes.TryGetValue(imeFajla, out instanca);
                    }
                    
                    if (instanca != null && DateTime.Now < instanca.Isticanje)
                    {
                        buffer = instanca.Data;
                        poruka = $"KES HIT ZA {imeFajla}";
                    }
                    else
                    {
                        lock (genLock)
                        {
                            int timeToLive = 500; //u milisekundama
                            lock (kesLock)
                            {
                                kes.TryGetValue(imeFajla , out instanca);
                            }
                            if (instanca != null)
                            {
                                if (DateTime.Now >= instanca.Isticanje)
                                {
                                    buffer = GenerisiXSLXfajl(csvFajl);
                                    lock (kesLock)
                                    {
                                        instanca.Data = buffer;
                                        instanca.Isticanje = DateTime.Now.AddMilliseconds(timeToLive);
                                    }
                                    poruka = $"OBNAVLJANJE KESA ZA {imeFajla}";
                                }
                                else
                                {
                                    buffer = instanca.Data;
                                    poruka = $"KES HIT ZA {imeFajla}";
                                }
                            }
                            else
                            {
                                buffer = GenerisiXSLXfajl(csvFajl);
                                lock (kesLock)
                                {
                                    kes.Add(imeFajla, new KesInstanca
                                    {
                                        Data = buffer,
                                        Isticanje = DateTime.Now.AddMilliseconds(timeToLive)
                                    });
                                    poruka = $"GENERISANJE I UPIS U KES ZA {imeFajla}";
                                }
                            }
                        }

                    }

                    Ispisi(poruka);

                    response.StatusCode = 200;
                    response.ContentLength64 = buffer.Length;
                    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    response.AddHeader("Content-Disposition", $"attachment; filename={imeFajla}.xlsx");

                    if (request.HttpMethod == "GET")
                        response.OutputStream.Write(buffer, 0, buffer.Length);
                    Ispisi($" DOBIJA {csvFajl} ZA {start.ElapsedMilliseconds} ms ");
                }
                catch(FileNotFoundException e)
                {
                    response.StatusCode = 404;
                    Ispisi($"NE POSTOJI FAJL {e.FileName} NA SERVERU");

                    using (var writer = new StreamWriter(response.OutputStream))
                    {
                        writer.Write($"Ne postoji fajl {e.FileName} na serveru");
                    }
                }
                catch (Exception ex)
                {
                    Ispisi($"GRESKA: {ex.Message}");
                    response.StatusCode = 500;
                }
                finally
                {
                    response.Close();
                }
            }
        }
        public byte[] GenerisiXSLXfajl(string imeFajla)
        {
            //Thread.Sleep(500);
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string path = Path.Combine(basePath, "Data");
            path = Path.Combine(path, imeFajla);

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
                return excelBytes.ToArray();
            }
        }
        public void Ispisi(string ispis)
        {
            lock (logLock)
            {
                Console.WriteLine($"{DateTime.Now:HH:mm:ss.fff} | " +
                                  $"Thread {Thread.CurrentThread.ManagedThreadId} | " +
                                  $"{ispis}");
            }
        }
        private void GracefulShutdown()
        {
            while (true)
            {
                string input = Console.ReadLine();
                if (input?.Trim().ToUpper() == "X")
                {
                    lock(logLock)
                    {
                        Console.WriteLine("GRACEFUL SHUTDOWN");
                    }
                    shutdown = true;
                    listener.Stop();
                    lock (queuelock)
                    {
                        Monitor.PulseAll(queuelock);
                    }
                    lock (obradaLock)
                    {
                        while (aktivneNiti > 0 || queue.Count > 0)
                        {
                            Monitor.Wait(obradaLock);
                        }
                    }
                    break;
                }
            }
        }
        private void ServerPreopterecen(HttpListenerContext context)
        {
            Ispisi("SERVER PREOPTERECEN");
            context.Response.StatusCode = 503;
            using (var writer = new StreamWriter(context.Response.OutputStream))
            {
                writer.Write("Server je preopterecen");
            }
            context.Response.Close();
        }
    }
}