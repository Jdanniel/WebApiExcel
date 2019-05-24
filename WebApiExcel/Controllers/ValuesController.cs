using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace WebApiExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        // GET api/values
        [HttpGet]
        public ActionResult Excel1()
        {
            var columnasEncabezado = new String[]
            {
                "AR",
                "NEGOCIO",
                "STATUS"
            };
            byte[] result;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Servicios");
                using (var cells = worksheet.Cells[1, 1, 1, 5])
                {
                    cells.Style.Font.Bold = true;
                }
                for (var i = 0; i < columnasEncabezado.Count(); i++)
                {
                    worksheet.Cells[1, i + 1].Value = columnasEncabezado[i];
                }

            }
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public ActionResult<string> Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
