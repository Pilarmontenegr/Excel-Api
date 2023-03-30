using Excel_Api.Rules;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic.FileIO;
using Excel_Api;

namespace Excel_Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [ApiKey]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("SueldoAnual")]
        
        public IActionResult ObtenerSueldoAnual(string nombre,int anual)
        {
            var sueldo1 = new Sueldo();
            sueldo1.MontoAnual = anual;
            sueldo1.Nombre = nombre;
           

            byte[] fileContent;
            fileContent = sueldo1.CrearExcel();

            string fileExtension;
            fileExtension = "xlsx";

            string fileContentType;
            fileContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            
            var fileName =$"SUELDO-API_{DateTime.Today:yyyy.MM.dd}";
            return File(fileContent,fileContentType,$"{fileName}.{fileExtension}");

            //return sueldo1;
            
        }

        //public Sueldo ObtenerSueldoAnual(string nombre, int anual)
        //{
        //    var sueldo1 = new Sueldo();


        //    sueldo1.MontoAnual = anual;
        //    sueldo1.Nombre = nombre;
        //    sueldo1.CrearExcel();
        //    return sueldo1;
        [HttpGet]
        [Route("SueldoMensual")]
        public IActionResult ObtenerSueldoMensual(int monto, string nombre)
        {
            var sueldo1 = new Sueldo();
            sueldo1.Nombre = nombre;
            sueldo1.Monto = monto;

            byte[] fileContent;
            fileContent = sueldo1.CrearExcel();

            string fileExtension;
            fileExtension = "xlsx";

            string fileContentType;
            fileContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


            var fileName = $"SUELDO-API_{DateTime.Today:yyyy.MM.dd}";
            return File(fileContent, fileContentType, $"{fileName}.{fileExtension}");

            


        }

    }
        
    
}