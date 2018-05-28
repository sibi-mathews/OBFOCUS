using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;

namespace OBFOCUS.UI.ServiceAccessor
{
    public static class HttpClientService
    {
        public static HttpClient WebApiClient = new HttpClient();

        static HttpClientService()
        {
            WebApiClient.BaseAddress = new Uri(ConfigurationManager.AppSettings["ApiURLHost"]);
            WebApiClient.DefaultRequestHeaders.Clear();
            WebApiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }
    }
}