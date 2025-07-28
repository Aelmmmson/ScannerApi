using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace ScannerApi
{
    public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            services.AddCors(options =>
            {
                options.AddPolicy("AllowFrontend", builder =>
                {
                    builder.WithOrigins(
                            "http://localhost:3000",
                            "http://localhost:8080",
                            "http://localhost:8081",
                            "http://localhost/my-app",
                            "http://localhost/vscanner",
                            "http://localhost" // Covers all subpaths
                        )
                        .AllowAnyMethod()
                        .AllowAnyHeader()
                        .AllowCredentials() // Support credentials (e.g., cookies, auth headers)
                        .SetIsOriginAllowedToAllowWildcardSubdomains(); // Handle subpaths
                });
            });
        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            // Comment out UseHttpsRedirection to avoid forcing HTTPS
            // app.UseHttpsRedirection();

            app.UseRouting();
            app.UseCors("AllowFrontend"); // Apply CORS before Authorization
            app.UseAuthorization();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}