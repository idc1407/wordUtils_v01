using System;

namespace WordUtilLib.Models
{
    public class WordUtilModel
    {

        public string SourceFileName { get; set; }
        public string TargetFileName { get; set; }
        public bool IsFooterTextChange { get; set; }
        public string FooterTextFind { get; set; }
        public string FooterTextReplace { get; set; }

        public bool IsHeaderTextChange { get; set; }
        public string HeaderTextFind { get; set; }
        public string HeaderTextReplace { get; set; }

        public bool IsBalanceSheetTableDelete { get; set; }
        public bool IsOtherOptionA { get; set; }
        public bool IsOtherOptionB { get; set; }
    }
}


