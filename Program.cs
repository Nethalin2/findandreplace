using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
           

            object fileName = Path.Combine("C:\\Users\\netha\\Documents\\FSharpTest\\FTEST", "ftestdoc.docx");

            Word.Application wordApp = new Word.Application { Visible = true };

            Word.Document aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

            aDoc.Activate();

            Word.Find fnd = wordApp.ActiveWindow.Selection.Find;

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = Word.WdFindWrap.wdFindContinue;

            fnd.Text = "test";
            fnd.Replacement.Text = "McTesty";

            fnd.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }
    }
}
