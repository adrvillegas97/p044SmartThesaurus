using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractTextWord
{
    class ExtractWord
    {
        private static ExtractWord instance;
        private Application wordApp;    // To use word
        private object missing;
        private Document document;

        /// <summary>
        /// 
        /// </summary>
        private ExtractWord()
        {
            // NO animation, NO visible
            wordApp = new Application();
            wordApp.ShowAnimation = false;
            wordApp.Visible = false;

            //Create a missing variable for missing value  
            missing = System.Type.Missing;
        }

        /// <summary>
        /// Accès au singleton
        /// </summary>
        /// <returns>l'instance singleton</returns>
        public static ExtractWord GetInstance()
        {
            if (instance == null)
            {
                instance = new ExtractWord();
            }
            return instance;
        }

        public void GetText(string path)
        {
            SetDocument(path);

            // To save temporaly the text from one paragraph
            string text = "";
            // To stock a string foreach paragraph
            List<string> textWord = new List<string>();

            // Get the document's paragraph
            for (int i = 0; i < document.Paragraphs.Count; ++i)
            {
                // Word doesn't have index 0, Trim to supr startings and endings spaces
                text = document.Paragraphs[i + 1].Range.Text.Trim();
                // Add the paragraph if not empty
                if (text != string.Empty)
                {
                    textWord.Add(text);
                }
            }

            CloseDocument();
        }

        private void SetDocument(string path)
        {
            document = wordApp.Documents.Open(path,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        private void CloseDocument()
        {
            ((_Document)document).Close();
            document = null;
        }

        public void CloseApp()
        {
            ((_Application)wordApp).Quit();
        }
    }
}
