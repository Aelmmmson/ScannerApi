using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using System;
using System.Threading;

namespace ScannerApi
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create a unique mutex for the application
            bool createdNew;
            using (var mutex = new Mutex(true, "ChequeScannerAppBackend", out createdNew))
            {
                if (!createdNew)
                {
                    Console.WriteLine("Another instance of the backend is already running. Please close it and try again.");
                    return;
                }

                try
                {
                    CreateHostBuilder(args).Build().Run();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fatal error starting application: {ex.Message}");
                    throw;
                }
            }
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder
                    .UseStartup<Startup>()
                    .UseUrls("http://localhost:5042"); // Forces the app to always bind to this port;
                });
    }
}
