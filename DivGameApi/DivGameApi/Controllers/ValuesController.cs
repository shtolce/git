using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DivGameApi.Models;
using DivGameApi.Services;
using Microsoft.Extensions.DependencyInjection;
using System.Runtime.Serialization.Json;
using Newtonsoft.Json;

namespace DivGameApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        Dialog1 dialog;
        public ValuesController(DialogService d)
        {
            dialog = d.dialogText;
        }
        // GET api/values
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public JsonResult Get(int id)
        {
            return new JsonResult(dialog);
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
