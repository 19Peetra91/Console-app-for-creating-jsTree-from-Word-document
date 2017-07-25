using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Novacode;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Upute
{  
    class Program
    {
        static void Main(string[] args)
        {
            var putanjadocx = ConfigurationManager.AppSettings["putanjadocx"].ToString();
            var izlazniDirZaHtml = ConfigurationManager.AppSettings["izlazniDirektorij"].ToString();

            List<IcUputeConverter.RootObject> uputa = new List<IcUputeConverter.RootObject>();

            foreach (var file in Directory.GetFiles(putanjadocx, "*.docx"))
            {
                var icUputa = new IcUputeConverter(file);
                uputa.Add(icUputa.root);
            }

            var output = JsonConvert.SerializeObject(uputa);
            var FileStringBuilder = new StringBuilder(output);
            string JsonString = (FileStringBuilder.Replace(@"\r\n", " ")).ToString();

            izlazniDirZaHtml += @"\tree.js";
            string js = ("var data='" + JsonString + "'").ToString();
            File.WriteAllText(izlazniDirZaHtml, js);

            //string[] filePaths = null;
            //string[] subDirectories = null;

            //filePaths = Directory.GetFiles(izlazniDirZaHtml, "*.*");
            //subDirectories = Directory.GetDirectories(izlazniDirZaHtml);

            //foreach (string subDir in subDirectories)
            //{
            //    IcHTMLconverter.PosaljiNaFtp(null, subDir);
            //}
            
            Console.WriteLine("Enter..");
            Console.ReadKey();
        }
    }
}