using HtmlAgilityPack;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Upute
{
    public class IcHTMLconverter
    {
        public static void pretvoriuHTML(string file, string IzlazniDirHTML)
        {
            if (file.Contains("%20"))
            {              
              file = file.Replace("%20", " ");
            }          

            Microsoft.Office.Interop.Word.Application worddoc = new Microsoft.Office.Interop.Word.Application();
            worddoc.Visible = false;
            
            var filepath = IzlazniDirHTML+= @"\"+ Path.GetFileNameWithoutExtension(file) + ".html";
            Console.WriteLine(Path.GetFileName(file));

            Document doc = worddoc.Documents.Open(FileName: file, ReadOnly: true);            
            doc.SaveAs(FileName: filepath, FileFormat: WdSaveFormat.wdFormatFilteredHTML);
            doc.Close();
            worddoc.Quit();
        }

        /// <summary>
        /// pronadji linkove ispravi ih u .html i pretvori ih iz docx u html document
        /// </summary>     
        //public static void FixLinkHtmlToc(HtmlDocument html, string putanjahtml)
        //{
        //    var linkovidocx = html.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name == "href" && s.Attributes.First().Value.Contains(".docx")).ToList();
        //    //var linkovidocx = html.DocumentNode.SelectNodes("//p[@class = 'MsoNormal' or @class = 'MsoBodyText2' or @class = 'MsoSubtitle' or @class = 'MsoListParagraphCxSpFirst' or @class = 'MsoListParagraphCxSpMiddle' or @class = 'MsoListParagraphCxSpLast']")
        //    //                 .Where(s => s.ChildNodes.First().Name == "a").Where(a=>a.Attributes["href"].Value.Contains(".docx")).ToList();

        //    foreach (HtmlNode node in linkovidocx)
        //    {
        //        var putanjadocx = ConfigurationManager.AppSettings["putanja"].ToString();

        //        string linkdocx = node.Attributes.First().Value;
        //        string linkzaSljedhtml = Path.GetFileName(Path.ChangeExtension(linkdocx, ".html"));
        //        string linkdocxzapretvaranje = Path.ChangeExtension(linkzaSljedhtml, ".docx");

        //        //pretvoriuHTML(putanjadocx += @"\" + linkdocxzapretvaranje, Path.GetDirectoryName(putanjahtml));

        //        var href = linkdocx.Replace(linkdocx, linkzaSljedhtml);

        //        node.SetAttributeValue("href", href);

        //        using (StringWriter writer = new StringWriter())
        //        {
        //            html.Save(writer);
        //            File.WriteAllText(putanjahtml, writer.ToString(), Encoding.UTF8);
        //        }
        //    }
        //}

        public static void FixLinkHtml(HtmlDocument html, string putanjahtml, string file)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(FileName: file, ReadOnly: true);
            Microsoft.Office.Interop.Word.Hyperlinks links = doc.Hyperlinks;

            if (links.Count > 0)
            {
                foreach (var l in links)
                {
                    var linksljeddocx = ((Microsoft.Office.Interop.Word.Hyperlink)l).Address;
                    if (linksljeddocx != null)
                    {
                        if (linksljeddocx.Contains(".docx"))
                        {
                            var paragraf = ((Microsoft.Office.Interop.Word.Hyperlink)l).Range.ParagraphStyle;
                            var ParagrafStyle = ((Microsoft.Office.Interop.Word.Style)paragraf).NameLocal;

                            //if (ParagrafStyle != "Heading 1" && ParagrafStyle != "Heading 2" && ParagrafStyle != "Heading 3" && ParagrafStyle != "Naslov 1" && ParagrafStyle != "Naslov 2" && ParagrafStyle != "Naslov 3")
                            //{
                                var putanjadocx = ConfigurationManager.AppSettings["putanja"].ToString();
                                //putanjadocx += @"\" + linksljeddocx;
                                var linkdocx = new StringBuilder(linksljeddocx);
                                linkdocx.Replace(" ", "%20");

                                if (linkdocx.ToString().Contains("../"))
                                {
                                    linkdocx.Replace("..", "");
                                }

                                var linkovidocx = html.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name == "href" && s.Attributes.First().Value.Contains("" + linkdocx + "")).ToList();

                                foreach (HtmlNode node in linkovidocx)
                                {
                                    string Linkdocx = node.Attributes.First().Value;
                                    string linkzaSljedhtml = Path.GetFileName(Path.ChangeExtension(Linkdocx, ".html"));

                                    //string linkdocxzapretvaranje = Path.ChangeExtension(linkzaSljedhtml, ".docx");                                     
                                    //pretvoriuHTML(putanjadocx += @"\" + linkdocxzapretvaranje, Path.GetDirectoryName(putanjahtml));

                                    var href = Linkdocx.Replace(Linkdocx, linkzaSljedhtml);

                                    node.SetAttributeValue("href", href);

                                    using (StringWriter writer = new StringWriter())
                                    {
                                        html.Save(writer);
                                        File.WriteAllText(putanjahtml, writer.ToString(), Encoding.UTF8);
                                    }

                                   if (ParagrafStyle != "Heading 1" && ParagrafStyle != "Heading 2" && ParagrafStyle != "Heading 3" && ParagrafStyle != "Naslov 1" && ParagrafStyle != "Naslov 2" && ParagrafStyle != "Naslov 3")
                                   {
                                    IcUputeConverter.Citajlinkove(putanjadocx += @"\" + linkdocx, Path.GetDirectoryName(putanjahtml));
                                   }
                                }
                            //}
                            //else
                            //{
                            //    var putanjadocx = ConfigurationManager.AppSettings["putanja"].ToString();
                            //    //putanjadocx += @"\" + linksljeddocx;
                            //    var linkdocx = new StringBuilder(linksljeddocx);
                            //    linkdocx.Replace(" ", "%20");

                            //    if (linkdocx.ToString().Contains("../"))
                            //    {
                            //        linkdocx.Replace("..", "");
                            //    }

                            //    var linkovidocx = html.DocumentNode.SelectNodes("//a").Where(s => s.Attributes.First().Name == "href" && s.Attributes.First().Value.Contains("" + linkdocx + "")).ToList();

                            //    foreach (HtmlNode node in linkovidocx)
                            //    {
                            //        string Linkdocx = node.Attributes.First().Value;
                            //        string linkzaSljedhtml = Path.GetFileName(Path.ChangeExtension(Linkdocx, ".html"));

                            //        //string linkdocxzapretvaranje = Path.ChangeExtension(linkzaSljedhtml, ".docx");                                     
                            //        //pretvoriuHTML(putanjadocx += @"\" + linkdocxzapretvaranje, Path.GetDirectoryName(putanjahtml));

                            //        var href = Linkdocx.Replace(Linkdocx, linkzaSljedhtml);

                            //        node.SetAttributeValue("href", href);

                            //        using (StringWriter writer = new StringWriter())
                            //        {
                            //            html.Save(writer);
                            //            File.WriteAllText(putanjahtml, writer.ToString(), Encoding.UTF8);
                            //        }
                            //    }
                            //}                      
                        }
                    }
                }
            }
            doc.Close();
            app.Quit();
        }

        public static void PosaljiNaFtp(string file)
        {
            Console.WriteLine("Šaljem na FTP!!");

            WebRequest ftpRequest = WebRequest.Create("ftp://ftp.ic-zupanja.hr/informaticki-zupanja/proba/" + Path.GetFileName(file));
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
            ftpRequest.Credentials = new NetworkCredential("iczupanja-001", "vgicX9ZE");

            StreamReader sourceStream = new StreamReader(file);
            //byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            byte[] fileData = File.ReadAllBytes(file);
            sourceStream.Close();
            //ftpRequest.ContentLength = fileContents.Length;
            ftpRequest.ContentLength = fileData.Length;
            
            Stream requestStream = ftpRequest.GetRequestStream();
            requestStream.Write(fileData, 0, fileData.Length);
            //requestStream.Write(fileContents, 0, fileContents.Length);

            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();

            Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);

            response.Close();
            

            //FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://ftp.ic-zupanja.hr/informaticki-zupanja/proba/"+ Path.GetFileName(file));
            //request.Method = WebRequestMethods.Ftp.UploadFile;

            //// This example assumes the FTP site uses anonymous logon.  
            //request.Credentials = new NetworkCredential("iczupanja-001", "vgicX9ZE");

            //// Copy the contents of the file to the request stream.  
            //StreamReader sourceStream = new StreamReader(file);
            //byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            //sourceStream.Close();
            //request.ContentLength = fileContents.Length;

            //Stream requestStream = request.GetRequestStream();
            //requestStream.Write(fileContents, 0, fileContents.Length);
            //requestStream.Close();

            //FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            //Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);

            //response.Close();

        }
                
    }
}