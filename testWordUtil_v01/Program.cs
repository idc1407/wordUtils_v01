using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordUtilLib;

namespace testWordUtil_v01
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string[] textReplce = { "Footer text goes here", "God is good" };

            await WordUtilLib.Main.Process(
                @"D:\itemp\temp1.docx",
                @"D:\itemp\temp2.docx",
                textReplce
                );
        
        
        }
    }
}
