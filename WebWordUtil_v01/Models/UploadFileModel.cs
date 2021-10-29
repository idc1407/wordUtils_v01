using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace WebWordUtil_v01.Models
{
    public class UploadFileModel
    {
        [Required(ErrorMessage = "Please Select a File.")]
        public HttpPostedFileBase File { get; set; }

        [DisplayName("Footer Text Change")]
        public bool IsFooterTextChange { get; set; }

        [DisplayName("Find Text")]
        [FooterTextFindValidation]
        public string FooterTextFind { get; set; }

        [DisplayName("Replace Text")]
        [FooterTextReplaceValidation]
        public string FooterTextReplace { get; set; }



        [DisplayName("Header Text Change")]
        public bool IsHeaderTextChange { get; set; }

        [DisplayName("Find Text")]
        [HeaderTextFindValidation]
        public string HeaderTextFind { get; set; }

        [DisplayName("Replace Text")]
        [HeaderTextReplaceValidation]
        public string HeaderTextReplace { get; set; }


        [DisplayName("Balance Sheet Table Delete")]
        public bool IsBalanceSheetTableDelete { get; set; }


        [DisplayName("Other Options A")]
        public bool IsOtherOptionA { get; set; }


        [DisplayName("Other Options B")]
        public bool IsOtherOptionB { get; set; }


    }


}


