



    $("body").on("click", "#btnAddRvwer", function () {


            //Get the reference of the Table's TBODY element.

            var tBody = $("#tblReviewers > TBODY")[0];

   

        var eml = $('#ddlEmployees').val();


        var f = searchtable(eml, 'tblReviewers');


            if (f != '') {

            alert('already added');
        return;

    }

    //Add Row.

    var row = tBody.insertRow(-1);

    //Add Name cell.

    var cell = $(row.insertCell(-1));


    cell.html($('#ddlEmployees :selected').text());


    //Add Country cell.

    cell = $(row.insertCell(-1));

    cell.html($('#ddlEmployees').val());

    //Add Button cell.

    cell = $(row.insertCell(-1));

            var btnRemoveRvr = $("<input />");

        btnRemoveRvr.attr("type", "button");

        btnRemoveRvr.attr("onclick", "Remove(this,'#tblReviewers');");

        btnRemoveRvr.val("Remove");

        cell.append(btnRemoveRvr);


    });


        $("body").on("click", "#btnAddVwer", function () {

            //Reference the Name and Country TextBoxes.


            var txtName = $("#txtName");

        var txtEmail = $("#txtEmail");

        var txtEmail = $('#ddlViewers').val();


        //To get the currently selected text:
        // $('#dropDownId :selected').text();



        //Get the reference of the Table's TBODY element.

        var tBody = $("#tblViewers > TBODY")[0];


        var eml = $('#ddlViewers').val();

        var f = searchtable(eml, 'tblViewers');


    
        // document.getElementById("ddlEmployees").contains(eml);

            if (f != '') {

            alert('already added');
        return;

    }

    //Add Row.

    var row = tBody.insertRow(-1);



    //Add Name cell.

    var cell = $(row.insertCell(-1));

    cell.html($('#ddlViewers :selected').text());


    //Add Country cell.

    cell = $(row.insertCell(-1));

    cell.html($('#ddlViewers').val());



    //Add Button cell.

    cell = $(row.insertCell(-1));

            var btnRemoveVwr = $("<input />");

        btnRemoveVwr.attr("type", "button");

        btnRemoveVwr.attr("onclick", "Remove(this,'#tblViewers');");

        btnRemoveVwr.val("Remove");

        cell.append(btnRemoveVwr);



        //Clear the TextBoxes.

        txtName.val("");


    });


        function searchtable(eml, tblname) {

            var tab = document.getElementById(tblname);
        var l = tab.rows.length;
        var s = '';
            for (var i = 0; i < l; i++) {
                var tr = tab.rows[i];
        var cll = tr.cells[1];

                if (cll.innerText.indexOf(eml) != -1) {
            //document.write(i + 1);

            s = 'found';
        break;
    }
}

return s;

}

         function Remove(button, tableid) {

            //Determine the reference of the Row using the Button.

            window.clearTimeout(window.timeoutHandle);
        window.timeoutHandle = window.setTimeout("SessionEnd()", window.sTimeout);

        var row = $(button).closest("TR");

        var name = $("TD", row).eq(0).html();

            if (confirm("Do you want to delete: " + name)) {

                //Get the reference of the Table.

                // var table = $("#tblReviewers")[0];

                var table = $(tableid)[0];



        //Delete the Table row using it's Index.

        table.deleteRow(row[0].rowIndex);

    }

};



        function DisplayProgressMessage(ctl, msg) {

            $(ctl).prop("disabled", true).text(msg);
        $(".submit-progress").removeClass("hidden");
        $("body").addClass("submit-progress-bg");
            setTimeout(function () {
            $(".submit-progress").removeClass("hidden");
        }, 1);

        return true;
    }


    function hasAlpha(str) {
        var code, i, len;

        for (i = 0, len = str.length; i < len; i++) {
            code = str.charCodeAt(i);
        if (code > 64 && code < 91)  // upper alpha (A-Z)
            return true;
    }
    return false;
};


    function isAlphaNumeric(str) {
        var code, i, len;

        for (i = 0, len = str.length; i < len; i++) {
            code = str.charCodeAt(i);
        if (!(code > 47 && code < 58) && // numeric (0-9)
                !(code > 64 && code < 91) && // upper alpha (A-Z)
                !(code > 96 && code < 123)) { // lower alpha (a-z)
                return false;
    }
}
return true;
};

function zeroPad(num, places) {
    var zero = places - num.toString().length + 1;
    return Array(+(zero > 0 && zero)).join("0") + num;
};


