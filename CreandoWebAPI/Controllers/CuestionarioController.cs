using CreandoWebApi.Entidades;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.AspNetCore.Hosting;

using Newtonsoft.Json.Linq;
using Syncfusion.EJ2.Navigations;
using Syncfusion.XlsIO;

using System.Dynamic;
using System.IO;
using System.Text;

namespace CreandoWebApi.Controllers
{

    [ApiController]
    [Route("api/cuestionario")]
    public class CuestionarioController : Controller
    {


        [HttpGet]

        public string ImportarExcel(string archivo)
        {
            ViewBag.Sheet1 = new TabHeader { Text = "Datos obligatorios" };
            ViewBag.Sheet2 = new TabHeader { Text = "Reactivos" };

     
            
                //Instantiate the spreadsheet creation engine.
                ExcelEngine excelEngine = new ExcelEngine();

                //Instantiate the Excel application object.
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the input Excel file
                FileStream stream = new FileStream(archivo, FileMode.Open, FileAccess.ReadWrite);
                IWorkbook book = application.Workbooks.Open(stream);
                stream.Close();

                //Access first worksheet
                IWorksheet worksheet = book.Worksheets[0];

                //Access a range
                IRange range = worksheet.Range["A1:H10"];

                MemoryStream jsonStream = new MemoryStream();

              
                book.SaveAsJson(jsonStream); //Save the entire workbook as a JSON stream
               

                excelEngine.Dispose();

                byte[] json = new byte[jsonStream.Length];

                //Read the Json stream and convert to a Json object
                jsonStream.Position = 0;
                jsonStream.Read(json, 0, (int)jsonStream.Length);
                string jsonString = Encoding.UTF8.GetString(json);
                    JObject jsonObject = JObject.Parse(jsonString);

               
                    //The first worksheet in the input document is converted to a JSON object and bind to the DataGrid in the first tab.
                    ViewBag.Tab1 = ((JArray)(jsonObject["Datos obligatorios"])).ToObject<List<CustomDynamicObject>>();

                    ////The second worksheet in the input document is converted to a JSON object and bind to the DataGrid in the second tab.
                    ViewBag.Tab2 = ((JArray)(jsonObject["Reactivos"])).ToObject<List<CustomDynamicObject>>();

           
            
                    
                    JArray DatosObligatorios = (JArray)jsonObject["Datos obligatorios"];
                    JArray Reactivos = (JArray)jsonObject["Reactivos"];
                    JArray Identificadores = (JArray)jsonObject["Identificadores"];

                      return jsonObject.ToString();
                
                

              
            
        }






















    }




    public class CustomDynamicObject : DynamicObject
    {
        /// <summary>
        /// The dictionary property used store the data
        /// </summary>
        internal Dictionary<string, object> properties = new Dictionary<string, object>();
        /// <summary>
        /// Provides the implementation for operations that get member values.
        /// </summary>
        /// <param name="binder">Get Member Binder object</param>
        /// <param name="result">The result of the get operation.</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = default(object);

            if (properties.ContainsKey(binder.Name))
            {
                result = properties[binder.Name];
                return true;
            }
            return false;
        }
        /// <summary>
        /// Provides the implementation for operations that set member values.
        /// </summary>
        /// <param name="binder">Set memeber binder object</param>
        /// <param name="value">The value to set to the member</param>
        /// <returns>true if the operation is successful; otherwise, false.</returns>
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            properties[binder.Name] = value;
            return true;
        }
        /// <summary>
        /// Return all dynamic member names
        /// </summary>
        /// <returns>the property name list</returns>
        public override IEnumerable<string> GetDynamicMemberNames()
        {
            return properties.Keys;
        }
    }












}
