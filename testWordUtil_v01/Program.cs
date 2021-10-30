using System.Threading.Tasks;

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
