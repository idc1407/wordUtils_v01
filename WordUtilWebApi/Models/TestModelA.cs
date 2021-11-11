using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace WordUtilWebApi.Models
{
    public class TestModelA
    {

        public string FirstName { get; set; }
        public IFormFile Image { get; set; }

    }
}
