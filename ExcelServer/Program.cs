using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelServer.Classes;
namespace ExcelServer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Server server = new Server();
            server.Pokreni();
        }
    }
}
