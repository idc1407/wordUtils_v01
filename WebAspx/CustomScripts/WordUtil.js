$(document).ready(function () {

    $('#File').change(function () {
        $("#smessage").hide();
        $("#emessage").hide();
    });

    $('#IsFooterTextChange').change(function () {
        console.log("I am here");
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


        $.post("https://localhost:5001/WeatherForecast", { id: 1, sourceFileName: "2pm" })
            .done(function (data) {
                alert("Data Loaded: " + data);
            });

    });


    $('#Button1').click(function () {
        //var data = { "id": 1, "SourceFileName": "test" };
        //$.ajax({
        //    headers: {
        //        'Accept': 'application/json',
        //        'Content-Type': 'application/json'
        //    },
        //    url: "https://localhost:5001/WeatherForecast",
        //    type: "POST",
        //    dataType: "json",
        //    data: JSON.stringify(data),
        //    success: function (data) {
        //        console.log(data);
        //    },
        //    failure: function (data) {
        //        console.log("failure");
        //        alert(data.responseText);
        //    },
        //    error: function (data) {
        //        console.log("error");
        //        alert(data.responseText);
        //    }
        //});


        var formData = new FormData();
        var fileInput = $('#FileUpload1')[0].files[0];

        formData.append("Image", fileInput);
        formData.append("fname", "all is good");


        $.ajax({
            url: "https://localhost:5001/WeatherForecast",
            type: 'POST',
            data: formData,
            processData: false,  
            contentType: false,
            success: function (result) {
                console.log("sucess");
            },
            error: function (jqXHR) {
            },
            complete: function (jqXHR, status) {
            }
        });

    });

});