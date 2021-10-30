using System;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace WebWordUtil_v01.Models
{
    public class WordUtilModel
    {
        [Required(ErrorMessage = "Please Select a File.")]
        public HttpPostedFileBase File { get; set; }

        [DisplayName("Footer Text Change")]
        public bool IsFooterTextChange { get; set; }

        [DisplayName("Find Text")]
        public string FooterTextFind { get; set; }

        [DisplayName("Replace Text")]
        public string FooterTextReplace { get; set; }



        [DisplayName("Header Text Change")]
        public bool IsHeaderTextChange { get; set; }

        [DisplayName("Find Text")]
        public string HeaderTextFind { get; set; }

        [DisplayName("Replace Text")]
        public string HeaderTextReplace { get; set; }


        [DisplayName("Balance Sheet Table Delete")]
        public bool IsBalanceSheetTableDelete { get; set; }


        [DisplayName("Other Options A")]
        public bool IsOtherOptionA { get; set; }


        [DisplayName("Other Options B")]
        public bool IsOtherOptionB { get; set; }

    }
}


