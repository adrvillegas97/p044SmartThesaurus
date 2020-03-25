using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace ExtractTextWord
{
    class Program
    {
        static void Main(string[] args)
        {
            // To use word : NO animation, NO visible
            Application wordApp = new Application();
            wordApp.ShowAnimation = false;
            wordApp.Visible = false;

            //Create a missing variable for missing value  
            object missing = System.Type.Missing;

            //Open a document  
            Document document = wordApp.Documents.Open(@"C:\Users\AdrCast\source\repos\ExtractTextWord\bin\Debug\Document1.docx",
            ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            // To save temporaly the text from one paragraph
            string text = "";
            // To stock a string foreach paragraph
            List<string> textWord = new List<string>();

            // Get the document's paragraph
            for (int i = 0; i < document.Paragraphs.Count; ++i)
            {
                // Word doesn't have index 0, Trim to supr startings and endings spaces
                text = document.Paragraphs[i+1].Range.Text.Trim();
                // Add the paragraph if not empty
                if (text != string.Empty)
                {
                    textWord.Add(text);
                }
            }
            
            // Close the document and close the app
            ((_Document)document).Close();
            ((_Application)wordApp).Quit();

            Console.ReadLine();
            // TODO : Convert in class and separete in methods
            //        Use the class like a singleton and implemented in factory
        }
    }
}
