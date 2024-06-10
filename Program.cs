using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using PostCuentasAlarmas.PostCuentas;
using PostCuentasAlarmas.PostParticiones;
using PostCuentasAlarmas.PostZonas;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;

using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace PostCuentasAlarmas
{
    public class Program
    {
        static async Task Main(string[] args)
        {
          //  await CrearCuenta();
          //   await CrearParticion();
           await CrearZonas();
        }

        public static async Task CrearCuenta()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //var urlLogin = "https://sap-api-qa.bsf.pe/api/Login";
            //var urlPost = "https://sap-api-qa.bsf.pe/api/AlarmAccount";    
            var urlLogin = "https://sap-api.bsf.pe/api/Login";
            var urlPost = "https://sap-api.bsf.pe/api/AlarmAccount";

            var cliente = new HttpClient();
            string jsonData = @"{
                            ""Usuario"": ""string"",
                            ""Password"": ""string""
                        }";
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
            var respuesta = await cliente.PostAsync(urlLogin, content);

            if (respuesta.IsSuccessStatusCode)
            {
                var lectura = await respuesta.Content.ReadAsStringAsync();
                var x = JsonConvert.DeserializeObject<Token>(lectura);
                string token = x.token.ToString();
                cliente.DefaultRequestHeaders.Authorization =  new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                List<Cuentas> cuentas = LeerCuentasDesdeExcel("C:\\Users\\ASANCHEZ\\Downloads\\ExcelCuentas\\CuentasExcel.xlsx"); // Reemplaza con la ruta real
                                                                                                                                  // Iterar sobre las cuentas y realizar la solicitud para cada una
                foreach (var cuenta in cuentas)
                {
                    string jsonDataCuenta = @"{
                                        ""U_BSF_ID"": """ + cuenta.U_BSF_ID + @""",
                                    ""Name"": """ + cuenta.Name + @"""
                                }";
                    var contentCuenta = new StringContent(jsonDataCuenta, Encoding.UTF8, "application/json");
                    var respuestaCuenta = await cliente.PostAsync(urlPost, contentCuenta);
                    if (respuestaCuenta.IsSuccessStatusCode)
                    {
                        var lecturaCuenta = await respuestaCuenta.Content.ReadAsStringAsync();
                        var cuentaResponse = JsonConvert.DeserializeObject<Cuentas>(lecturaCuenta);
                        Console.WriteLine(cuentaResponse);
                    }
                    else
                    {
                        Console.WriteLine($"Error al procesar la cuenta {cuenta}: {respuestaCuenta.StatusCode}");
                    }
                }
                Console.WriteLine(cuentas.Count());
                Console.ReadLine();

            }
            Console.ReadLine();
        }

        public static async Task CrearParticion()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var urlLogin = "https://sap-api.bsf.pe/api/Login";
            var urlPost = "https://sap-api.bsf.pe/api/AlarmPartition";

            var cliente = new HttpClient();
            string jsonData = @"{
                            ""Usuario"": ""string"",
                            ""Password"": ""string""
                        }";
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
            var respuesta = await cliente.PostAsync(urlLogin, content);

            if (respuesta.IsSuccessStatusCode)
            {
                var lectura = await respuesta.Content.ReadAsStringAsync();
                var x = JsonConvert.DeserializeObject<Token>(lectura);
                string token = x.token.ToString();
                cliente.DefaultRequestHeaders.Authorization =  new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                List<Particion> cuentas = LeerCuentasParticionesDesdeExcel("C:\\Users\\ASANCHEZ\\Downloads\\ExcelCuentas\\CuentaParticion.xlsx"); // Reemplaza con la ruta real
                                                                                                                                                  // Iterar sobre las cuentas y realizar la solicitud para cada una
                foreach (var cuenta in cuentas)
                {
                    string jsonDataCuenta = @"{
""Code"" : """ + cuenta.Code + @""",
""U_BSF_ID"": """ + cuenta.U_BSF_ID + @""",
""U_BSF_NOM"": """ + cuenta.U_BSF_NOM + @"""

                                } ";

                    var contentCuenta = new StringContent(jsonDataCuenta, Encoding.UTF8, "application/json");
                    var respuestaCuenta = await cliente.PostAsync(urlPost, contentCuenta);
                    if (respuestaCuenta.IsSuccessStatusCode)
                    {
                        var lecturaCuenta = await respuestaCuenta.Content.ReadAsStringAsync();
                        var cuentaResponse = JsonConvert.DeserializeObject<Particion>(lecturaCuenta);
                        Console.WriteLine(cuentaResponse);
                    }
                    else
                    {
                        Console.WriteLine($"Error al procesar la cuenta {cuenta}: {respuestaCuenta.StatusCode}");
                        var lecturaError = await respuestaCuenta.Content.ReadAsStringAsync();
                        Console.WriteLine(lecturaError);
                    }
                }
                Console.WriteLine(cuentas.Count());
                Console.ReadLine();
            }
        }

        public static async Task CrearZonas()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var urlLogin = "https://sap-api.bsf.pe/api/Login";
            var urlPost = "https://sap-api.bsf.pe/api/AlarmZone";

            var cliente = new HttpClient();
            string jsonData = @"{
                            ""Usuario"": ""string"",
                            ""Password"": ""string""
                        }";
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
            var respuesta = await cliente.PostAsync(urlLogin, content);

            if (respuesta.IsSuccessStatusCode)
            {
                var lectura = await respuesta.Content.ReadAsStringAsync();
                var x = JsonConvert.DeserializeObject<Token>(lectura);
                string token = x.token.ToString();
                cliente.DefaultRequestHeaders.Authorization =  new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                List<Zonas> cuentas = LeerCuentasZonasDesdeExcel("C:\\Users\\ASANCHEZ\\Downloads\\ExcelCuentas\\CuentaZonas.xlsx"); // Reemplaza con la ruta real
                                                                                                                                                  // Iterar sobre las cuentas y realizar la solicitud para cada una
                foreach (var cuenta in cuentas)
                {
                    string jsonDataCuenta = @"{
""Code"" : """ + cuenta.Code + @""",
""U_BSF_ID"": """ + cuenta.U_BSF_ID + @""",
""U_BSF_NOM"": """ + cuenta.U_BSF_NOM + @""",
""U_BSF_PARTICION"": """ + cuenta.U_BSF_PARTICION + @"""

                                } ";

                    var contentCuenta = new StringContent(jsonDataCuenta, Encoding.UTF8, "application/json");
                    var respuestaCuenta = await cliente.PostAsync(urlPost, contentCuenta);
                    if (respuestaCuenta.IsSuccessStatusCode)
                    {
                        var lecturaCuenta = await respuestaCuenta.Content.ReadAsStringAsync();
                        var cuentaResponse = JsonConvert.DeserializeObject<Zonas>(lecturaCuenta);
                        Console.WriteLine(cuentaResponse);
                    }
                    else
                    {
                        Console.WriteLine($"Error al procesar la cuenta {cuenta}: {respuestaCuenta.StatusCode}");
                        var lecturaError = await respuestaCuenta.Content.ReadAsStringAsync();
                        Console.WriteLine(lecturaError);
                    }
                }
                Console.WriteLine(cuentas.Count());
                Console.ReadLine();
            }
        }
        static List<Cuentas> LeerCuentasDesdeExcel(string rutaArchivo)
        {
            // Lee las cuentas desde el archivo Excel y las devuelve como una lista de objetos Cuenta
            List<Cuentas> cuentas = new List<Cuentas>();

            try
            {
                FileInfo file = new FileInfo(rutaArchivo);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount+1; row++) // Start from row 2 to skip header
                    {
                        string U_BSF_ID = worksheet.Cells[row, 1].Value.ToString().Trim();
                        string Name = worksheet.Cells[row, 2].Value.ToString().Trim();

                        cuentas.Add(new Cuentas { U_BSF_ID = U_BSF_ID, Name = Name });
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error al leer el archivo Excel: {e.Message}");
            }

            return cuentas;
        }


        static List<Particion> LeerCuentasParticionesDesdeExcel(string rutaArchivo)
        {
            // Lee las cuentas desde el archivo Excel y las devuelve como una lista de objetos Cuenta
            List<Particion> cuentas = new List<Particion>();

            try
            {
                FileInfo file = new FileInfo(rutaArchivo);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount+1; row++) // Start from row 2 to skip header
                    {
                        string Code = worksheet.Cells[row, 1].Value.ToString().Trim();
                        string U_BSF_ID = worksheet.Cells[row, 2].Value.ToString().Trim();
                        string U_BSF_NOM = worksheet.Cells[row, 3].Value.ToString().Trim();

                        cuentas.Add(new Particion { Code = Code, U_BSF_ID = U_BSF_ID, U_BSF_NOM = U_BSF_NOM });
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error al leer el archivo Excel: {e.Message}");
            }

            return cuentas;
        }




        static List<Zonas> LeerCuentasZonasDesdeExcel(string rutaArchivo)
        {
            // Lee las cuentas desde el archivo Excel y las devuelve como una lista de objetos Cuenta
            List<Zonas> cuentas = new List<Zonas>();

            try
            {
                FileInfo file = new FileInfo(rutaArchivo);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount+1; row++) // Start from row 2 to skip header
                    {
                        string Code = worksheet.Cells[row, 1].Value.ToString().Trim();
                        string U_BSF_ID = worksheet.Cells[row, 2].Value.ToString().Trim();
                        string U_BSF_NOM = worksheet.Cells[row, 3].Value.ToString().Trim();
                        string U_BSF_PARTICION = worksheet.Cells[row, 4].Value.ToString().Trim();

                        cuentas.Add(new Zonas { Code = Code, U_BSF_ID = U_BSF_ID, U_BSF_NOM = U_BSF_NOM, U_BSF_PARTICION = U_BSF_PARTICION });
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error al leer el archivo Excel: {e.Message}");
            }

            return cuentas;
        }


    }


}
