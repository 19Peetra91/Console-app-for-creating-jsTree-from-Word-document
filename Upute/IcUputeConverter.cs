using HtmlAgilityPack;
using Microsoft.Office.Interop.Word;
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

namespace Upute
{
    public class IcUputeConverter
    {
        public int i = 0;
        public List<RootObject> tree = new List<RootObject>();
        public RootObject root = new RootObject();
        public class AAttr
        {
            public string href { get; set; }
        }     
        public class Child
        {
            public string text { get; set; }
            public AAttr a_attr = new AAttr();
            public List<Child> children = new List<Child>();
        }           
        public class RootObject
        {
            public string text { get; set; }
            public List<Child> children = new List<Child>();
        }

        public IcUputeConverter(string file = null)
        {
            //var tree = new List<Uputa>();
            //var tree = new List<RootObject>();
            //var root = new RootObject();
            root.text = Path.GetFileNameWithoutExtension(file);

            var izlazniDirZaHtml = ConfigurationManager.AppSettings["izlazniDirektorij"].ToString();
                  
            IcHTMLconverter.pretvoriuHTML(file, izlazniDirZaHtml);
            
            //ucitaj pretvoreni html
            var putanjahtml = izlazniDirZaHtml;
            putanjahtml += @"\";
            var adresaHtml = Path.ChangeExtension(Path.GetFileName(file), ".html");
            HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();
            html.Load(putanjahtml += adresaHtml);

            //popravi linkove u html-u
            IcHTMLconverter.FixLinkHtml(html, putanjahtml, file);

            //trazi TOC
            var Toc = html.DocumentNode.Descendants("p").Where(s=> s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc1") || s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc2") || s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc3")).ToList();
            
            if (Toc.Count > 0)
            {
                foreach (var t in Toc)
                {
                    HtmlAgilityPack.HtmlDocument htmlInner = new HtmlAgilityPack.HtmlDocument();
                    htmlInner.LoadHtml(t.InnerHtml);

                    var HtmlTag_a = htmlInner.DocumentNode.SelectSingleNode("//a");

                    var TocId = HtmlTag_a.Attributes.First().Value.Replace("#", " ").TrimStart();
                    var htmlnodovi = html.DocumentNode.SelectNodes("//a").Where(s =>s.Attributes.First().Value == TocId).FirstOrDefault();

                    //var tt = new Uputa();
                    //tt.text = t.InnerText;
                    //tt.href = putanjahtml + "#" + TocId;

                    var tt = new Child();

                    //tt.text = t.InnerText;
                    tt.a_attr.href = adresaHtml + "#" + TocId;
                                    

                    if (t.Attributes.First().Value == "MsoToc1")
                    {
                        var txt = htmlInner.DocumentNode.SelectSingleNode("//a").InnerText;
                        var onlyLetters = new String(txt.Where(c => Char.IsLetter(c) || Char.IsWhiteSpace(c)).ToArray());
                        
                        tt.text = onlyLetters;
                        
                        //tt.text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[1]").InnerText;

                        root.children.Add(tt);
                    }
                    else if (t.Attributes.First().Value == "MsoToc2")
                    {
                        //var text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[1]").InnerText.ToString();
                        var text = htmlInner.DocumentNode.SelectSingleNode("//a").InnerText.ToString();
                        var onlyLetters = new String(text.Where(c => Char.IsLetter(c) || Char.IsWhiteSpace(c)).ToArray());

                        tt.text = onlyLetters;
                         
                        var parent = root.children.Last();                       
                        parent.children.Add(tt);
                    }
                    else if (t.Attributes.First().Value == "MsoToc3")
                    {
                        var djed = root.children.Last();
                        var otac = djed.children.Last();

                        var naslov3 = htmlInner.DocumentNode.SelectSingleNode("//a").InnerText;
                        var onlyLetters = new String(naslov3.Where(c => Char.IsLetter(c) || Char.IsWhiteSpace(c)).ToArray());

                        tt.text = onlyLetters;
                        //tt.text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[1]").InnerText.ToString();
                        otac.children.Add(tt);
                    }
                    
                    HtmlAgilityPack.HtmlDocument htmlH1 = new HtmlAgilityPack.HtmlDocument();
                    htmlH1.LoadHtml(htmlnodovi.ParentNode.InnerHtml);

                   var HtmlTag_s_linkom = htmlH1.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name.Equals("href")).SingleOrDefault();

                    if (HtmlTag_s_linkom != null)
                    {
                        //izvuci link iz Toc-a za sljed docx
                        var linkSljedDocx = Path.GetFileName(Path.ChangeExtension(HtmlTag_s_linkom.Attributes.First().Value, ".docx"));

                        //prepravi linkove u TOC-u iz docx u html
                        //IcHTMLconverter.FixLinkuTOCu(htmlnodovi, html, putanjahtml);

                        //izradi fullpath za sljed docx
                        //string putanjalinkdocx = Path.GetDirectoryName(file);
                        var putanjadocx = ConfigurationManager.AppSettings["putanja"].ToString();

                        putanjadocx += @"\" + linkSljedDocx;
                        var linkdocx = new StringBuilder(putanjadocx);
                        linkdocx.Replace("%20", " ");

                        napunichildren(linkdocx.ToString(), izlazniDirZaHtml, root, tt.text);
                    }
                }

                tree.Add(root);
            }

            //var output = JsonConvert.SerializeObject(tree);
            //var FileStringBuilder = new StringBuilder(output);
            //string JsonString = (FileStringBuilder.Replace(@"\r\n", " ")).ToString();

            //izlazniDirZaHtml += @"\tree.js";
            //string js = ("var data='" + JsonString + "'").ToString();
            //File.WriteAllText(izlazniDirZaHtml, js);
        }

        void napunichildren(string file, string izlazniDirHtml, RootObject jobject, string parent = null, string parentToc3 = null)
        {
            //Pretvori u HTML
            IcHTMLconverter.pretvoriuHTML(file, izlazniDirHtml);

           // izradi link za pretvoreni html
            var putanjahtml = izlazniDirHtml;
            putanjahtml += @"\";
            var adresaHtml = Path.ChangeExtension(Path.GetFileName(file), ".html");
            HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();
            html.Load(putanjahtml += adresaHtml);

            IcHTMLconverter.FixLinkHtml(html, putanjahtml, file);
            //var Toc = html.DocumentNode.SelectNodes("//p").Where(s => s.Attributes.First().Name.Equals("class") && s.Attributes.First().Value.Equals("MsoToc1") || s.Attributes.First().Value.Equals("MsoToc2") || s.Attributes.First().Value.Equals("MsoToc3")).ToList();
            var Toc = html.DocumentNode.Descendants("p").Where(s => s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc1") || s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc2") || s.Attributes.Contains("class") && s.Attributes["class"].Value.Contains("MsoToc3")).ToList();

            if (Toc.Count > 0)
            {
                //bool sw_jednom = true;

                foreach (var h in Toc)
                {
                    //if (sw_jednom)
                    //{
                    //    //up.children = new List<Uputa>();
                    //    jobject.children = new List<Child2>();
                    //    sw_jednom = false;
                    //}
                    HtmlAgilityPack.HtmlDocument htmlInner = new HtmlAgilityPack.HtmlDocument();
                    htmlInner.LoadHtml(h.InnerHtml);

                    var HtmlTag_a = htmlInner.DocumentNode.SelectSingleNode("//a");

                    var TocId = HtmlTag_a.Attributes.First().Value.Replace("#", " ").TrimStart();
                    var htmlnodovi = html.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Value == TocId).FirstOrDefault();
                    
                     //var parent1 = jobject.children.Last();
                    //var parent1 = jobject.Find(s => s.text == parent);         
                
                    var tt = new Child();
                    //tt.text = h.InnerText;
                    tt.a_attr.href = adresaHtml + "#" + TocId;

                    //VrtiKrozListu(jobject.children, parent, tt, h.Attributes.First().Value);

                    if (h.Attributes.First().Value == "MsoToc1")
                    {
                        tt.text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[2]").InnerText.ToString();

                        VrtiKrozListu(jobject.children, parent, tt);

                        parent = tt.text;
                    }
                    else if (h.Attributes.First().Value == "MsoToc2")
                    {

                        //tt.text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[2]").InnerText.ToString();

                        for (i = 1; i < 4; i++)
                        {
                            var text = htmlInner.DocumentNode.SelectSingleNode("//a/text()[" + i + "]").InnerText.ToString();
                            if (!text.Contains(".") && text != null)
                            {
                                tt.text = text;
                                i = 0;
                                break;
                            }
                        }
                        VrtiKrozListu(jobject.children, parent, tt);
                        parentToc3 = tt.text;                                                                                         
                    }
                    else if (h.Attributes.First().Value == "MsoToc3")
                    {
                        var naslov3 = htmlInner.DocumentNode.SelectSingleNode("//a/text()[1]");
                        if (naslov3 == null)
                        {
                            tt.text = h.InnerText;
                        }
                        else
                        {
                            tt.text = naslov3.InnerText;
                        }

                        VrtiKrozListu(jobject.children, null, tt, parentToc3);
                    }
                    //tt.href = putanjahtml + "#" + TocId;                                        
                    //parent.children.Add(tt);

                    HtmlAgilityPack.HtmlDocument htmlH1 = new HtmlAgilityPack.HtmlDocument();
                    htmlH1.LoadHtml(htmlnodovi.ParentNode.InnerHtml);

                    var HtmlTag_s_linkom = htmlH1.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name.Equals("href")).SingleOrDefault();

                    if (HtmlTag_s_linkom != null)
                    {
                        //izvuci link it Toc-a za sljed docx
                        //var linkSljedDocx = Path.GetFileName(HtmlTag_s_linkom.Attributes.First().Value);
                        var linkSljedDocx = Path.GetFileName(Path.ChangeExtension(HtmlTag_s_linkom.Attributes.First().Value, ".docx"));

                        //izradi fullpath za sljed docx
                        string putanjalinkdocx = Path.GetDirectoryName(file);
                        putanjalinkdocx += @"\" + linkSljedDocx;
                        var linkdocx = new StringBuilder(putanjalinkdocx);
                        linkdocx.Replace("%20", " ");
                        
                        napunichildren(linkdocx.ToString(), izlazniDirHtml, jobject, tt.text);
                    }                    
                }          
            }                      
        }

        public void VrtiKrozListu(List<Child> children, string parent, Child tt, string parentToc3 = null)
        {
            foreach (var c in children)
            {
                if (parent != null)
                {
                    if (c.text == parent)
                    {
                        c.children.Add(tt);
                        break;
                    }
                }
                else
                {
                    if (c.text == parentToc3)
                    {
                        c.children.Add(tt);
                        break;
                    }
                }

                VrtiKrozListu(c.children, parent, tt, parentToc3);                
            }
        }
       
        //citaj linkove kroz docx pretvori ih u html 
        public static void Citajlinkove(string file, string izDirHtml)
        {
            if (file.Contains("../"))
            {
                file = file.Replace("../", "");
            }
            file = file.Replace("%20", " ");
            //file = file.Replace("/", @"\");

            IcHTMLconverter.pretvoriuHTML(file, izDirHtml);

            //ucitaj pretvoreni html
            var putanjahtml = izDirHtml;
            putanjahtml += @"\";
            var adresaHtml = Path.ChangeExtension(Path.GetFileName(file), ".html");
            HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();
            html.Load(putanjahtml += adresaHtml);

            var linkovidocx = html.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name == "href" && s.Attributes.First().Value.Contains(".docx")).ToList();
            //var linkovidocx = html.DocumentNode.SelectNodes("//p[@class = 'MsoNormal' or @class = 'MsoBodyText2' or @class = 'MsoSubtitle' or @class = 'MsoListParagraphCxSpFirst' or @class = 'MsoListParagraphCxSpMiddle' or @class = 'MsoListParagraphCxSpLast']")
            //                 .Where(s => s.ChildNodes.First().Name == "a").Where(a=>a.Attributes["href"].Value.Contains(".docx")).ToList();

            foreach (HtmlNode node in linkovidocx)
            {
                var putanjadocx = ConfigurationManager.AppSettings["putanja"].ToString();

                string linkdocx = node.Attributes.First().Value;
                string linkzaSljedhtml = Path.GetFileName(Path.ChangeExtension(linkdocx, ".html"));
                string linkdocxzapretvaranje = Path.ChangeExtension(linkzaSljedhtml, ".docx");
                
                //IcHTMLconverter.pretvoriuHTML(putanjadocx += @"\" + linkdocxzapretvaranje, Path.GetDirectoryName(putanjahtml));

                var href = linkdocx.Replace(linkdocx, linkzaSljedhtml);

                node.SetAttributeValue("href", href);

                using (StringWriter writer = new StringWriter())
                {
                    html.Save(writer);
                    File.WriteAllText(putanjahtml, writer.ToString(), Encoding.UTF8);
                }

                Citajlinkove(putanjadocx += @"\" + linkdocxzapretvaranje, Path.GetDirectoryName(putanjahtml));
            }           
        }
                       
        //public void children(string file, string parent = null, JObject jobject = null)
        //{
        //    var izlazniDirektorij = ConfigurationManager.AppSettings["izlazniDirektorij"].ToString();

        //    //Pretvori u HTML
        //    HtmlConverterHelper.PretvoriUHtml(file, izlazniDirektorij);

        //    var putanjahtml = Path.GetFullPath(izlazniDirektorij);
        //    putanjahtml += @"\";

        //    var adresaHtml = Path.ChangeExtension(Path.GetFileName(file), ".html");

        //    HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();
        //    html.Load(putanjahtml += adresaHtml);

        //    //var toc1 = html.DocumentNode.SelectNodes("//p[@class = \"pt-TOC1\" or @class = \"pt-Sadraj1\"]/a");
        //    var toc = html.DocumentNode.SelectNodes("//p[@class = \"pt-TOC1\" or @class = \"pt-TOC2\" or @class = \"pt-Sadraj1\" or @class = \"pt-Sadraj2\"]/a");
        //    //var toc2 = html.DocumentNode.SelectNodes("//p[@class = \"pt-TOC2\" or @class = \"pt-Sadraj2\"]/a");
        //    //var toc3 = html.DocumentNode.SelectNodes("//p[@class = \"pt-TOC3\" or @class = \"pt-Sadraj3\"]/a");

        //    if (toc != null)
        //    {
        //        foreach (var t in toc)
        //        {                    
        //            lista.Add(new IcUputeConverter { Link = putanjahtml, Heading = t.InnerText });                                
        //        }

        //        if (jobject == null)
        //        {
        //            JObject noviobject = new JObject(
        //                new JProperty("text", "Upute"),
        //              new JProperty("state",
        //              new JObject(
        //                  new JProperty("opened", true),
        //                  new JProperty("selected", true))),
        //              new JProperty("children",
        //              new JArray(
        //                   new JObject(
        //                       new JProperty("text", parent),
        //                       new JProperty("children",
        //                       new JArray(
        //                           from x in lista
        //                           select new JObject(
        //                               new JProperty("text", x.Heading),
        //                               new JProperty("a_attr", new JObject(
        //                                   new JProperty("href", putanjahtml))))))))));

        //            jobject = noviobject;
        //            string a = jobject.ToString();

        //            lista.Clear();
        //      }
        //       else
        //      {
        //            JToken noviobject =
        //               new JObject(
        //                   new JProperty("text", parent),
        //                    new JProperty("children",
        //                    new JArray(
        //                        from qu in lista
        //                        select new JObject(
        //                            new JProperty("text", qu.Heading),
        //                            new JProperty("a_attr", new JObject(
        //                                new JProperty("href", putanjahtml)))))));
                    
        //            var json = JObject.Parse(jobject.ToString());
        //            var children = json["children"][0];

        //            var childzaPromjenu = children.SelectToken("$.children[?(@.text == '"+ parent +"')]");

        //            childzaPromjenu.Replace(noviobject);
                   
        //            string gg = json.ToString();
        //        }
        //    }

        //    var FileStringBuilder = new StringBuilder(file);
        //    FileStringBuilder.Replace("%20", " ");
        //    var filedocx = FileStringBuilder.ToString();

        //    DocX docx = DocX.Load(file);

        //    var linkovi = docx.Hyperlinks.ToList();
        //    if (linkovi != null)
        //    {
        //        foreach (var l in linkovi)
        //        {
        //            if (l.Uri != null)
        //            {
        //                string putanjalinkdocx = Path.GetDirectoryName(file);
        //                putanjalinkdocx += @"\";
        //                putanjalinkdocx += l.Uri;
        //                var Builder = new StringBuilder(putanjalinkdocx);
        //                Builder.Replace("%20", " ");

        //                file = Builder.ToString();

        //                children(file, l.Text, jobject);
        //            }
        //        }
        //    }

        //    string putanjatree = ConfigurationManager.AppSettings["izlazniDirektorij"].ToString();
        //    putanjatree += @"\tree.json";

        //    string output = Newtonsoft.Json.JsonConvert.SerializeObject(jobject, Newtonsoft.Json.Formatting.Indented);

        //    File.WriteAllText(putanjatree, output);                    
        //}
    }
}