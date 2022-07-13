function downloadUrl(url) {
    window.open(url, '_self');
}

let button = document.querySelector("#upload")
let Download = document.querySelector("#download")
var result = {};
let db2


button.addEventListener("click", function() {
    upload()
})


// upload excel file
function upload() {
    var files = document.getElementById('file_upload').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

//Excel to json
function excelFileToJSON(file) {
    try {

        var reader = new FileReader();
        reader.readAsBinaryString(file);



        reader.onload = function(e) {

            var data = e.target.result;

            var workbook = XLSX.read(data, {
                type: 'binary', cellDates: true, dateNF:"jj/mm/aaaa  hh:mm:ss "
            });

            console.log("workbook = ", workbook)

            result = {};

            workbook.SheetNames.forEach(function(sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                    console.log("roa = ", roa)
                    console.log("result {} =", result)
                }
            });
            //displaying the json result
            var resultEle = document.getElementById("json-result");
            resultEle.value = JSON.stringify(result, null, 4);
            resultEle.style.display = 'block';
            console.log("Json ", result);
            //console.log("stringy = ", JSON.stringify(result, null, 4))
            compteur_sinistre_total(result.Sheet1)
        }



    } catch (e) {
        console.error(e);
    }
}

//-------------------------------------------------------------------------------
let db = []

function compteur_sinistre_total(arry) {
var n,zero,cent,adeter;
let personne

    for (let i = 0; i < arry.length; i++) {
        n = 0
        zero = 0
        cq = 0
        cent = 0
        adeter = 0
        const element = arry[i];

        for (let j = 0; j < arry.length; j++) {
            const element2 = arry[j];
            if (element.CONDUCTEUR == element2.CONDUCTEUR) {
                if (element2['% RESP.'] == 0) {
                    zero++
                }
                if (element2['% RESP.'] == 50) {
                    cq++
                }
                if (element2['% RESP.'] == 100) {
                    cent++
                }
                if (element2['% RESP.'] == 150) {
                    adeter++
                }
                n++
            }
        }
        //console.log(element.CONDUCTEUR,n)
        //console.log(element.CONDUCTEUR,zero,cq,cent,adeter)
        personne = {
            SITE: element.SITE,
            CONDUCTEUR: element.CONDUCTEUR,
            "LITIGE TOTAL" : n,
            "RESP 0%": zero,
            "RESP 50%": cq,
            "RESP 100%": cent,
            "RESP A DEF": adeter,
            TYPE: element.TYPE
        }
        
        
        db.push(personne)
    }
    
    let doublon = "doublon"
    let personne_total = 0


    for (let k = 0; k < db.length; k++) {
        
        let occurence = 0;
        const elementary = db[k];
        for (let j = 0; j < arry.length; j++) {
            
            const elemento = arry[j];
            if (elementary.CONDUCTEUR == elemento.CONDUCTEUR) {
                if (occurence != 0) {
                    //db.splice(j,elementary['LITIGE TOTAL'])
                    db[j]="doublon"
                }
                //console.log(db[k].CONDUCTEUR,occurence)
                occurence++
            }
        }
        db2 = db.filter(item => item !== doublon)
        
    }

    let restotal = 0

    for (let l = 0; l < db2.length; l++) {
        const elementr = db2[l];
        if (elementr != doublon) {
            personne_total++
        }
        restotal += elementr['LITIGE TOTAL']
    }

    console.log("db trier ",db2,personne_total,restotal)


    var resultEle2 = document.getElementById("json-result2");
    var h3 = document.querySelectorAll("h3")
    h3.forEach(element => {
        element.style.display = 'block'
    });
    resultEle2.value = JSON.stringify(db2, null, 4);
    resultEle2.style.display = 'block';
    console.log("Json db2", db2);
    result['Sheet2'] = db2
}

//-------------------------------------------------------------------------------


var xlsRows;
var xlsHeader;

var xlsRows2;
var xlsHeader2;

Download.addEventListener("click", function() {

    var createXLSLFormatObj = [];
    var createXLSLFormatObj2 = [];

    /* XLS1 Head Columns */
    xlsHeader = ["Nº INTERNE", "DATE CREATION", "DERNIERE RELANCE", "STATUT", "N° SEM", "DATE", "DECLARATION", "REF DOSSIER", "IMMAT", "SITE", "ACHAT/LOC", "CONDUCTEUR", "TYPE", "% RESP.", "PRIME ACC", "TIERS 1", "MONTANT HT SUM: 167650.62", "MONTANT REMB. SUM: 23228.66", "MT TIERS PROV SUM: 114946.19", "MT TIERS REEL SUM: 124926.34", "ANCIENNETE - MOIS", "AGENCE INTERIM", "SITUATION ACCIDENT", "TEL PORTABLE", "TEL DOMICILE", "CLIENT", "NOM PDV", "CAMERA", "COM", "NOM GARAGE", "ADRESSE GARAGE", "VILLE G", "TEL G"];
    xlsHeader2 = ["SITE","CONDUCTEUR","LITIGE TOTAL","RESP 0%","RESP 50%","RESP 100%","RESP A DEF"];

    /* XLS1 Rows Data */
    xlsRows = result
    console.log("Resultat = ", xlsRows)
   
    createXLSLFormatObj.push(xlsHeader);
    console.log("createXLSLFormatObj = ", createXLSLFormatObj)

    createXLSLFormatObj2.push(xlsHeader2);
    console.log("createXLSLFormatObj2 = ", createXLSLFormatObj2)

    xlsRows.Sheet1.forEach(element => {
        var innerRowData = [];
        console.log("element : ", element)

        xlsHeader.forEach(val => {
            innerRowData.push(element[val]);
            console.log("valeur : ", val)
        });
        createXLSLFormatObj.push(innerRowData);
    });

    xlsRows.Sheet2.forEach(element => {
        var innerRowData2 = [];
        console.log("element : ", element)

        xlsHeader2.forEach(val => {
            innerRowData2.push(element[val]);
            console.log("valeur : ", val)
        });
        createXLSLFormatObj2.push(innerRowData2);
    });

    // for (let i = 0; i < xlsRows.Sheet1.length; i++) {

    //     const element = xlsRows.Sheet1[i];
    //     console.log("for ",element)
    //     var innerRowData = [];
    //     innerRowData.push(element)
    //     createXLSLFormatObj.push(innerRowData);
    //     console.log("createXLSLFormatObj for = ",createXLSLFormatObj)
    // }




    // $.each(xlsRows, function(index, value) {
    //     var innerRowData = [];
    //     $("tbody").append('<tr><td>' + value.EmployeeID + '</td><td>' + value.FullName + '</td></tr>');
    //     $.each(value, function(ind, val) {
    //         innerRowData.push(val);
    //     });
    //     createXLSLFormatObj.push(innerRowData);
    // });


    /* File Name */
    var filename = document.getElementById("file_upload").files[0].name + ".xlsx"

    /* Sheet Name */
    var ws_name = "Feuille complète";

    var ws_name2 = "Feuille résumé"

    if (typeof console !== 'undefined') console.log(new Date());

    var wb = XLSX.utils.book_new(),
        ws = XLSX.utils.aoa_to_sheet(createXLSLFormatObj),
        ws2 = XLSX.utils.aoa_to_sheet(createXLSLFormatObj2);

    /* Add worksheet to workbook */
    XLSX.utils.book_append_sheet(wb, ws, ws_name);
    XLSX.utils.book_append_sheet(wb, ws2, ws_name2);

    /* Write workbook and Download */
    if (typeof console !== 'undefined') console.log(new Date());
    XLSX.writeFile(wb, filename);
    //XLSX.writeFile(wb2, filename);
    if (typeof console !== 'undefined') console.log(new Date());

})