using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System.Xml.Linq;

namespace SAP_XRISK_WS
{
    public class Initializer
    {
      internal  static SAPbobsCOM.Company _Company;
        public static SAPbobsCOM.Company Company
        {
            get
            {
                if(_Company==null || !_Company.Connected) { 
                string salt = "!@#908!112aFgKts";
                var path1 = Path.Combine(HostingEnvironment.MapPath("~/App_Data"), "cnx.p");
                string text = System.IO.File.ReadAllText(path1);
                text = text.DecryptMe(salt);
                XDocument doc = XDocument.Parse(text);
                _Company = new SAPbobsCOM.Company();
                _Company.Server = doc.Descendants("Server").First().Value;
                _Company.LicenseServer = doc.Descendants("LicenseServer").First().Value;
                _Company.DbServerType = (SAPbobsCOM.BoDataServerTypes)Enum.Parse(typeof(SAPbobsCOM.BoDataServerTypes), doc.Descendants("DbServerType").First().Value);
                _Company.DbUserName = doc.Descendants("DbUserName").First().Value;
                _Company.DbPassword = doc.Descendants("DbPassword").First().Value;
                _Company.CompanyDB = doc.Descendants("CompanyDB").First().Value;
                _Company.UserName = doc.Descendants("UserName").First().Value;
                _Company.Password = doc.Descendants("Password").First().Value;
                var i=_Company.Connect();

                }
                return _Company;
            }
        }
        public static void AppInitialize()
        {
            var c = Company;

            // This will get called on startup
        }
    }
}