using Microsoft.AspNetCore.Mvc;
using PROYECTO_EXCEL_CORE.Models;
using System.Diagnostics;

using System.Data;
using System.Data.SqlClient;


using ClosedXML.Excel;

namespace PROYECTO_EXCEL_CORE.Controllers
{
    public class HomeController : Controller
    {
        private readonly string cadenaSQL;

        public HomeController(IConfiguration config)
        {
            cadenaSQL = config.GetConnectionString("cadenaSQL");
        }

        public IActionResult Index()
        {
            return View();
        }

        //USAR REFERENCIAS ===> SYSTE.DATA
        public IActionResult Exportar_Excel(string fechaInicio, string fechaFin)
        {
            DataTable tabla_cliente = new DataTable();

            //=========== PRIMERO - OBTENER EL DATA ADAPTER ===========
            using (var conexion = new SqlConnection(cadenaSQL)) { 
                conexion.Open();
                using (var adapter = new SqlDataAdapter()) {

                    adapter.SelectCommand = new SqlCommand("sp_reporte_cliente", conexion);
                    adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaInicio", fechaInicio);
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaFin", fechaFin);
                   
                   

                    adapter.Fill(tabla_cliente);
                }
            }


            //usar referencias
            //=========== SEGUNDO - INSTALAR ClosedXML ===========
            using (var libro = new XLWorkbook()) {

                tabla_cliente.TableName = "Clientes";
                var hoja = libro.Worksheets.Add(tabla_cliente);
                hoja.ColumnsUsed().AdjustToContents();

                using (var memoria = new MemoryStream()) {

                    libro.SaveAs(memoria);

                    var nombreExcel = string.Concat("Reporte ", DateTime.Now.ToString(), ".xlsx");

                    return File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                }
            }



        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}