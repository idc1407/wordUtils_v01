using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace WebWordUtil_v01.Models
{
    public class FooterTextFindValidation : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var uploadFileModel = (UploadFileModel)validationContext.ObjectInstance;
            if (uploadFileModel.IsFooterTextChange && String.IsNullOrEmpty(uploadFileModel.FooterTextFind)) 
            {
                return new ValidationResult("Footer Find Text is Required");
            }
            else
            {
                return ValidationResult.Success;
            }
        }
    }


    public class FooterTextReplaceValidation : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var uploadFileModel = (UploadFileModel)validationContext.ObjectInstance;
            if (uploadFileModel.IsFooterTextChange && String.IsNullOrEmpty(uploadFileModel.FooterTextReplace))
            {
                return new ValidationResult("Footer Replace Text is Required");
            }
            else
            {
                return ValidationResult.Success;
            }
        }
    }





    public class HeaderTextFindValidation : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var uploadFileModel = (UploadFileModel)validationContext.ObjectInstance;
            if (uploadFileModel.IsHeaderTextChange && String.IsNullOrEmpty(uploadFileModel.HeaderTextFind))
            {
                return new ValidationResult("Header Find Text is Required");
            }
            else
            {
                return ValidationResult.Success;
            }
        }
    }


    public class HeaderTextReplaceValidation : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var uploadFileModel = (UploadFileModel)validationContext.ObjectInstance;
            if (uploadFileModel.IsHeaderTextChange && String.IsNullOrEmpty(uploadFileModel.HeaderTextReplace))
            {
                return new ValidationResult("Header Replace Text is Required");
            }
            else
            {
                return ValidationResult.Success;
            }
        }
    }



}