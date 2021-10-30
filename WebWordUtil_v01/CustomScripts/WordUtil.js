$(document).ready(function () {

    $('#File').change(function () {
        $("#smessage").hide();
        $("#emessage").hide();
    });

    $('#IsFooterTextChange').change(function () {
        if (this.checked) {
            $('#FooterTextFind').attr('required', true);
            $('#FooterTextReplace').attr('required', true);
        }
        else {
            $('#FooterTextFind').attr('required', false);
            $('#FooterTextFind').removeClass('field-validation-error').next('span[data-valmsg-for]').removeClass("field-validation-error").addClass("field-validation-valid").html("");

            $('#FooterTextReplace').attr('required', false);
            $('#FooterTextReplace').removeClass('field-validation-error').next('span[data-valmsg-for]').removeClass("field-validation-error").addClass("field-validation-valid").html("");
        }
    });

    $('#IsHeaderTextChange').change(function () {
        if (this.checked) {
            $('#HeaderTextFind').attr('required', true);
            $('#HeaderTextReplace').attr('required', true);
        }
        else {
            $('#HeaderTextFind').attr('required', false);
            $('#HeaderTextFind').removeClass('field-validation-error').next('span[data-valmsg-for]').removeClass("field-validation-error").addClass("field-validation-valid").html("");

            $('#HeaderTextReplace').attr('required', false);
            $('#HeaderTextReplace').removeClass('field-validation-error').next('span[data-valmsg-for]').removeClass("field-validation-error").addClass("field-validation-valid").html("");
        }
    });


    $('#process').click(function () {
        if ($("#wform").valid()) {
            $("#loading").show();
        }
    });

});