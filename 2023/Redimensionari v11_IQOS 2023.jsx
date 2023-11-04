app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.generalPreferences.pageNumbering = PageNumberingOptions.absolute;

app.scriptPreferences.enableRedraw = false;
app.layoutWindows[0].transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myDocument = app.documents[0];
var myPage_length = app.documents[0].pages.length
var myPages = app.documents[0].pages

myDocument.layers.everyItem().locked = false;
myDocument.guides.everyItem().remove();

///*************** TEST LAYERE VIZIBIL ID SI HW ************************************************************/
try {
    // app.documents[0].layers.item('Layer 1').name = 'art';
    app.documents[0].layers.add({ name: "vizibil", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers.add({ name: "id", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers.itemByName("vizibil").move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers.itemByName("id").move(LocationOptions.AT_BEGINNING);

} catch (e) {
    // alert(e);
    // app.documents[0].layers.itemByName("vizibil").move(LocationOptions.AT_BEGINNING);
    // app.documents[0].layers.itemByName("id").move(LocationOptions.AT_BEGINNING);

}




// try {
//     

//     app.documents[0].layers.itemByName("HW").move(LocationOptions.AFTER, app.documents[0].layers.itemByName("vizibil"));
//     // app.documents[0].layers.add({ name: 'HW', layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);


// } catch (e) {
//     // alert(e);
//     alert("HW layer NU exista!" + "\n" + "Document will close");
//     app.documents[0].close(SaveOptions.no);
//     // return;
//     exit();
//     // app.documents[0].layers.itemByName("HW").move(LocationOptions.AT_BEGINNING);

// }
///*************** TEST LAYERE VIZIBIL ID SI HW  END************************************************************/


var myFile_calea = app.documents[0].filePath;
// alert("calea e "+myFile_calea)	
var myFileName_full = app.documents[0].fullName + "";
var myFileName = app.documents[0].name + "";

var myFileName_full_length = myFileName_full.length
var myFileName_length = myFileName.length

var myFileName0 = myFileName.substr(0, myFileName.lastIndexOf("."));
// alert("myFileName0 "+myFileName0);



///*************** ORDONEAZA PAGINI DUPA RATIE START   ************************************************************/


myDocument.documentPreferences.facingPages = false;
var spread = app.activeDocument.spreads.everyItem()
spread.allowPageShuffle = true;
// myDocument.documentPreferences.allowSpreadShuffle = true;
myDocument.documentPreferences.allowPageShuffle = true;
// myDocument.documentPreferences.preserveLayoutWhenShuffling = true;


var myRatiaUnu = new Array;


citestepaginile(myRatiaUnu)

function citestepaginile(myRatiaUnu) {

    for (i = 0; i < app.documents[0].pages.length; i++) {

        myPage = app.documents[0].pages.item(i);

        var b = myPage.bounds;
        var W_ = b[3] - b[1]; // 
        var H_ = b[2] - b[0]; // 
        var ratia = W_ / H_;
        myRatiaUnu.push(ratia);

    }

    // alert("ratia din for din functia citeste " + myRatia)
    // alert("ratia length " + myRatia.length)
    comparaRatia_simuta(myRatiaUnu)
}


// alert("ratia din afara functiei " + myRatia)
// alert("ratia length " + myRatia.length)


function comparaRatia_simuta(myRatiaUnu) {
    for (var myCounter = 0; myCounter < (myRatiaUnu.length - 1); myCounter++) {



        if (myRatiaUnu[myCounter] > myRatiaUnu[(myCounter + 1)]) {
            // alert("counter din if " + myCounter)
            // alert("ratia din compara " + myRatia)
            // alert("ratia length din compara " + myRatia.length)
            // alert("counter " + myCounter + " counter plus 1  " + Number(myCounter + 1))
            // alert("ratia counter " + myRatia[myCounter] + " ratia Anticounter  " + myRatia[myCounter + 1])


            app.documents[0].spreads.item(myCounter).move(LocationOptions.AFTER, app.documents[0].spreads.item(myCounter + 1));

            myRatiaUnu = [];

            citestepaginile(myRatiaUnu)
        }


    }
}


app.documents[0].save();

///*************** ORDONEAZA PAGINI DUPA RATIE END    ************************************************************/





var definitionsFile = File(myFile_calea + "/" + myFileName0 + ".txt");
// alert("definitionsFile "+definitionsFile);

definitionsFile.open("r");
var countLines_l = 0;
while (!definitionsFile.eof) {
    var numarLinii = countLines_l++;
    var readedLine_l = definitionsFile.readln().split("\t"); // umplutura ca altfel da eroare
}
// alert("JOB line " + numarLinii);




////************************************************************ RATIA MASTER  */
var myPageNames = new Array;
for (myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
    myPageNames.push(app.documents[0].pages.item(myCounter).name);
}

// alert(" "+myPageNames)

var myRatia = new Array;
for (myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
    myPage = app.documents[0].pages.item(myCounter);
    var b = myPage.bounds;
    var W_ = b[3] - b[1]; // 
    var H_ = b[2] - b[0]; // 
    var ratia = W_ / H_;
    myRatia.push(ratia);
}

////************************************************************ RATIA MASTER END */






// ==================================================START PROGRESS BAR
var progressWin = CreateProgressBar();
progressWin.show();
progressWin.update(); // poate merge pe windows
progressWin.pb.minvalue = 0;
progressWin.pb.maxvalue = countLines_l;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
CreateProgressBar()
function CreateProgressBar() {
    var w = new Window("window", "Redimensionare");
    w.pb = w.add("progressbar", [12, 12, 1250, 24], 0, undefined);
    w.st = w.add("statictext");
    w.st.bounds = [0, 0, 1200, 20];
    w.st.alignment = "left";
    return w;
}
// ================================================== END PROGRESS BAR 1


definitionsFile.close("r");
definitionsFile.open("r");
var countLines = -1;
while (!definitionsFile.eof) {
    countLines++;

    //************************************************************************************************************* START ------ CITESTE DOC TEXT  */


    var readedLine = definitionsFile.readln().split("\t");//split("\t");
    // alert("readedLine[0] : "+readedLine[0]+"\n"+"readedLine[1] : "+readedLine[1]+"\n"+"readedLine[2] : "+readedLine[2]+"\n"+"readedLine[3] : "+
    // readedLine[3]+"\n"+"readedLine[4] : "+readedLine[4]+"\n"+"readedLine[5] : "+readedLine[5])
    if (countLines != 0) {
        // ==================================================  PROGRESS BAR
        var finalFileName = readedLine[7];
        var varInfoCol4 = readedLine[5];
        var varComment = readedLine[6];
        // var finalFileName = readedLine[7];
        var idNumar = readedLine[0]
        progressWin.pb.value = (countLines + 1);
        progressWin.st.text = "Processing file - " + finalFileName + "  (" + (countLines + 0) + " / " + (countLines_l - 1) + ")";
        progressWin.update();
        // ==================================================  PROGRESS BAR END


        // ==================================================  CITESTE W SI H


        var primu_W_mm = (readedLine[1]);
        var primu_W_mm = primu_W_mm.replace(/\,/g, ".");
        var primuH_mm = (readedLine[2]);
        var primuH_mm = primuH_mm.replace(/\,/g, ".");
        var aldoilea_W_mm = (readedLine[3]);
        var aldoilea_W_mm = aldoilea_W_mm.replace(/\,/g, ".");
        var aldoilea_H_mm = (readedLine[4]);
        var aldoilea_H_mm = aldoilea_H_mm.replace(/\,/g, ".");


        // alert("idNumar "+idNumar+"\n"+"vizibil_W_mm "+vizibil_W_mm+"\n"+"vizibil_H_mm "+vizibil_H_mm+"\n"+"total_W_mm "+total_W_mm+"\n"+"total_H_mm "+total_H_mm)
        var varAria_viz = primu_W_mm * primuH_mm
        var varAria_total = aldoilea_W_mm * aldoilea_H_mm

        if (varAria_viz < varAria_total) {
            var vizibil_W = Number(primu_W_mm * 2.83464567); //2.83464567
            var vizibil_H = Number(primuH_mm * 2.83464567);
            var total_W = Number(aldoilea_W_mm * 2.83464567);
            var total_H = Number(aldoilea_H_mm * 2.83464567);
            var vizibil_W_mm = readedLine[1];
            var vizibil_H_mm = readedLine[2];
            var total_W_mm = readedLine[3];
            var total_H_mm = readedLine[4];
        } else {
            var vizibil_W = Number(aldoilea_W_mm * 2.83464567);
            var vizibil_H = Number(aldoilea_H_mm * 2.83464567);
            var total_W = Number(primu_W_mm * 2.83464567);
            var total_H = Number(primuH_mm * 2.83464567);
            var vizibil_W_mm = readedLine[3];
            var vizibil_H_mm = readedLine[4];
            var total_W_mm = readedLine[1];
            var total_H_mm = readedLine[2];
        }
        // alert("total_H "+total_H)
        // alert("se executa "+countLines)

        // ==================================================  CITESTE W SI H END


        var ratia_final = vizibil_W / vizibil_H;

        ///// ********************************************************** RATIA FINALA PT UN DOCUMENT DECLINAT DIN MASTER

        var myDocument = app.documents[0];
        var myPage_length = app.documents[0].pages.length
        var myPages = app.documents[0].pages

        // alert("pagDeExtras: " + pagDeExtras)
        var ratiaMaster = new Array;
        ratiaMaster = myRatia;

        /////.......................>>>>>>>>>>>>>>>>>>>>>>>>>>
        // var ratiaMaster = [0.34, 0.54, 0.63, 1, 1.48, 1.77, 2.58, 3.3, 4, 5.17];
        // var xxtiaMaster = [0.00, 1.00, 2.00, 3, 4.00, 5.00, 6.00, 7.0, 8, 9.0];
        // var ratiNewFile = [0.335, 0.489311164, 0.5, 0.69023569, 1, 1.39047619, 2, 2.847619048, 3.214285714, 4.69047619, 9.380952381]
        // var ratia_final = 1.4


        var closestElement = findClosestElement(myRatia, ratia_final); // !! myRatia este ratia din master

        var pagDeExtras = closestElement;
        //// alerta pagini de extras
        // alert("pagDeExtras: " + pagDeExtras)

        function findClosestElement(arr, target) {
            // var closest = arr[0];
            var closest = 0;
            var minDifference = Math.abs(target - arr[0]);

            for (p = 1; p < arr.length; p++) {
                var difference = Math.abs(target - arr[p]);
                if (difference < minDifference) {
                    minDifference = difference;
                    // closest = arr[p];
                    closest = p;
                }
            }

            return closest;
        }






        /////.......................>>>>>>>>>>>>>>>>>>>>>>>>>>

        function dummyFunctionOldFind() { //// dummy

            //ratia_final
            // myRatia
            // myCounter
            // var pagDeExtras = myCounter;
            // var text_myRatia = myRatia[myCounter]
            // var myRatia_myCounter = parseFloat(text_myRatia);

            for (var myCounter = 0; myCounter < myRatia.length; myCounter++) {
                var val_jumatea_intervalului = (myRatia[myCounter + 1] - myRatia[myCounter]) / 2
                var text_myRatia = myRatia[myCounter]
                var myRatia_myCounter = parseFloat(text_myRatia);
                var jumate_interval = val_jumatea_intervalului + myRatia_myCounter
                // alert("val juamte "+val_jumatea_intervalului+"  juamte int "+jumate_interval)
                // alert("1 counter e " + myCounter + "\n" + " ratia_final " + ratia_final + "\n" + " myRatia[myCounter + 1] " + myRatia[myCounter + 1] +
                //     "\n" + " jumatea_intervalului!!: " + jumate_interval)
                if (ratia_final > myRatia[myCounter]) {
                    // alert("2 " + "ratia_final > myRatia[myCounter]")
                    // ratia noastra cautata e mai mare sau egala cu capatul de jos al intervalului;
                    // in continuare verificam daca suntem pe ultimul element din array (daca intervalul are si un capat superior sau nu)
                    if (myCounter == myRatia.length - 1) {
                        // alert("3 " + "myCounter == myRatia.length - 1")
                        // caz de exceptie, suntem pe ultimul interval (am ajuns la capatul array-ului);
                        // limita superioara a intervalului este infinit, nu mai avem ce verifica, am gasit intervalul
                        var pagDeExtras = myCounter;
                        break; // ne-am gasit intervalul, iesi din "for"
                    } else if (ratia_final <= myRatia[myCounter + 1]) {
                        // alert("4 " + "ratia_final <= myRatia[myCounter + 1]")
                        // cazul normal, suntem intre 2 intervale, si ratia noastra e strict mai mica decat limita superioara
                        // comparam daca e mai aproape de primul element sau mai aproape de ultimul element din interval

                        if (ratia_final <= jumate_interval) {
                            // alert("5 ratia_final " + ratia_final + "jumate_interval " + jumate_interval)
                            var pagDeExtras = myCounter;
                            // alert("4b counter e " + myCounter + "pag de extras este " + pagDeExtras + "\n" + " ratia_final " + ratia_final +
                            //     "\n" + " myRatia[myCounter + 1] " + myRatia[myCounter + 1] +
                            //     "\n" + " jumatea_intervalului " + jumate_interval)
                            break;
                        }
                        else if (ratia_final >= jumate_interval) {
                            // alert("6 ratia_final " + ratia_final + "jumate_interval " + jumate_interval)
                            var pagDeExtras = myCounter + 1;
                            // alert("6b pag de extras este " + pagDeExtras + " counter e " + myCounter)
                            break;
                        }
                        // break; // ne-am gasit intervalul, iesi din "for"
                    }
                } else if (ratia_final <= myRatia[0]) {
                    // alert("7 " + ratia_final <= myRatia[0])
                    var pagDeExtras = 0;
                    break;
                }

            }

        }/// dummy function
        // alert("pag de extras are nr " + pagDeExtras)
        /////***************************************************************cauta ratia end */



        ////****************************************************************** extract page */
        var myPageNames = new Array;
        for (myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
            myPageNames.push(app.documents[0].pages.item(myCounter).name);
        }

        // alert(myPage_length)

        // var pagDeExtras = 1 //************!!! */
        for (var i = myPage_length - 1; i >= 0; i--) {



            if (i > pagDeExtras) { //|| i > pagDeExtras
                // alert("i mai mare este " + i)

                myPages[i].remove();
            }
            if (i < pagDeExtras) {

                // alert(i)
                // alert("i mai mic este " + i)
                myPages[i].remove();

            }
        }

        var numarRatieRotunjit_u = Number(myRatia[pagDeExtras])
        var numarRatieRotunjit = numarRatieRotunjit_u.toFixed(3);
        var numeRatieFolder1 = numarRatieRotunjit.toString();
        var numeRatieFolder = numeRatieFolder1.replace(/\./g, "_");
        // alert("numeRatieFolder " + numeRatieFolder)
        var myFile_calea = app.documents[0].filePath;
        // alert("calea e "+myFile_calea)	
        var myFileName_full = app.documents[0].fullName + "";
        var myFileName = app.documents[0].name + "";

        var myFileName_full_length = myFileName_full.length
        var myFileName_length = myFileName.length

        // creare foldere cu ratia in nume
        // var f = new Folder(myFile_calea + "/" + ("_ratia_" + (numeRatieFolder)));
        // f.create();

        // creare foldere cu nume result


        try {

            var f = new Folder(myFile_calea + "/" + ("_result"));
            f.create();

        } catch (e) {
            // alert(e);


        }

        var finalFileName = finalFileName.replace(/\./g, "_");
        var finalFileName = finalFileName.replace(/\,/g, "_");
        var finalFileName = finalFileName.replace(/\:/g, "_");
        var finalFileName = finalFileName.replace(/\;/g, "_");
        var finalFileName = finalFileName.replace(/`/g, "_");
        var finalFileName = finalFileName.replace(/\!/g, "_");
        var finalFileName = finalFileName.replace(/\?/g, "_");
        var finalFileName = finalFileName.replace(/\>/g, "_");
        var finalFileName = finalFileName.replace(/\</g, "_");
        var finalFileName = finalFileName.replace(/\//g, "_");
        var finalFileName = finalFileName.replace(/\[/g, "_");
        var finalFileName = finalFileName.replace(/\\/g, "_");
        var finalFileName = finalFileName.replace(/\|/g, "_");
        var finalFileName = finalFileName.replace(/\]/g, "_");
        var finalFileName = finalFileName.replace(/\{/g, "_");
        var finalFileName = finalFileName.replace(/\}/g, "_");
        var finalFileName = finalFileName.replace(/\*/g, "_");
        var finalFileName = finalFileName.replace(/\^/g, "_");
        var finalFileName = finalFileName.replace(/\$/g, "_");
        var finalFileName = finalFileName.replace(/\&/g, "_");
        var finalFileName = finalFileName.replace(/\"/g, "_");
        var finalFileName = finalFileName.replace(/\'/g, "_");

        //var finalFileName = finalFileName.replace(/[.,:;`!?>/<\[\\|}\]*^$&"'']/g, "_");



        // *************************************************************************************** MAI JOS SE FAC OPERATII PE PAGINA */

        // operatii
        scaleDocument(vizibil_W, vizibil_H, total_W, total_H);

        resizeDocument(total_W, total_H);

        idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar, varComment, varInfoCol4);

        aliniereHwDocument(vizibil_W, vizibil_H, total_W, total_H);

        aliniereElemente(vizibil_W, vizibil_H, total_W, total_H);

        lockLayers();

        // ************************************************************************************************************ MAI SUS SE FAC OPERATII PE PAGINA */
        var savedocname = finalFileName + "_.indd";
        // var savedoc = myDocument.save(File(myFile_calea + "/" + ("_ratia_" + (numeRatieFolder)) + "/" + savedocname));
        var savedoc = myDocument.save(File(myFile_calea + "/" + "_result" + "/" + savedocname));
        savedoc.close(SaveOptions.no);
        // alert("document close")

        ////************************************************************************************************************* extract page END */


        ////************************************************************************************************************* OPEN MASTER AGAIN  */
        app.open(File(myFile_calea + "/" + myFileName), false);
        // var myDocument = app.open(File("/c/myTestDocument.indd"), false);!!!



    }
    //************************************************ END ------ CITESTE DOC TEXT  */


}

//***************************************************************************** START ------ FUNCTII



function scaleDocument(vizibil_W, vizibil_H, total_W, total_H) {
    // alert("idNumar "+idNumar+"\n"+"vizibil_W_mm "+vizibil_W_mm+"\n"+"vizibil_H_mm "+vizibil_H_mm+"\n"+"total_W_mm "+total_W_mm+"\n"+"total_H_mm "+total_H_mm)
    // alert("scaleDocument")

    app.documents[0].zeroPoint = [0, 0];
    app.documents[0].viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    var pages = app.documents[0].pages;
    pages[0].marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
    pages[0].layoutRule = LayoutRuleOptions.SCALE;
    // pages[0].layoutRule = LayoutRuleOptions.OFF; 
    pages[0].resize(CoordinateSpaces.SPREAD_COORDINATES,
        AnchorPoint.CENTER_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [vizibil_W, vizibil_H])
    ///***** */

}



function resizeDocument(total_W, total_H) {
    // alert("idNumar "+idNumar+"\n"+"vizibil_W_mm "+vizibil_W_mm+"\n"+"vizibil_H_mm "+vizibil_H_mm+"\n"+"total_W_mm "+total_W_mm+"\n"+"total_H_mm "+total_H_mm)
    // alert("resizeDocument")

    pages = app.documents[0].pages ///app.activeDocument
    pages[0].marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
    //pages[0].layoutRule = LayoutRuleOptions.SCALE;
    pages[0].layoutRule = LayoutRuleOptions.OFF;
    pages[0].resize(CoordinateSpaces.SPREAD_COORDINATES,
        AnchorPoint.CENTER_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [total_W, total_H])



}



function idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar, varComment) {



    // expandToPageBounds(bleed);

    /************* */
    // total_H
    // total_W
    // vizibil_H
    // vizibil_W
    var m_top = (total_H - vizibil_H) / 2;
    var m_left = (total_W - vizibil_W) / 2;
    var m_right = m_left;
    var m_bottom = m_top;

    // m_topq = m_top / 2.83464567;
    // m_leftq = m_left / 2.83464567;
    // alert(" m_top "+m_topq+" m_left "+m_leftq);


    var myPage = app.documents[0].pages.item(0);
    myPage.marginPreferences.properties = {
        top: m_top,
        left: m_left,
        right: m_right,
        bottom: m_bottom
    };

    app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
    var myItem_vizibil = myPage.rectangles.add({ geometricBounds: [m_top, m_left, m_top + vizibil_H, m_left + vizibil_W] });

    var myColorforvizibil = app.documents[0].colors.add({ name: "C=0 M=0 Y=100 K=0", colorValue: [0, 0, 100, 0], model: ColorModel.PROCESS, space: ColorSpace.CMYK });
    myItem_vizibil.label = "Vizibil";
    myItem_vizibil.fillColor = "None";
    myItem_vizibil.strokeColor = myColorforvizibil;
    myItem_vizibil.strokeWeight = 1;
    myItem_vizibil.strokeTint = 100;
    myItem_vizibil.strokeAlignment = StrokeAlignment.INSIDE_ALIGNMENT;

    // myRectangle_ptvizibil = app.selection[0];
    //    var currentfilename= app.activeDocument
    //             vizibil_W_q=vizibil_W/2.8346438836889;
    // vizibil_H_q=vizibil_H/2.8346438836889;
    // total_Wq=total_W/2.8346438836889;
    // total_Hq=total_H/2.8346438836889;
    //     alert(" w vizibil pt "+" este "+vizibil_W_q+" h vizibil pt "+vizibil_H_q+" w t pt "+" este "+total_Wq+" h t pt "+total_Hq);

    /************* */

}


function aliniereHwDocument(vizibil_W, vizibil_H, total_W, total_H) {



    // app.documents[0].activeLayer = app.documents[0].layers.itemByName("HW");
    var myPage = app.documents[0].pages.item(0);
    var b_pgebounds = myPage.bounds;

    var w_t = b_pgebounds[3] - b_pgebounds[1];
    var h_t = b_pgebounds[2] - b_pgebounds[0];
    // alert("w_t "+w_t+"\n"+"h_t "+h_t+"\n"+"idNumar "+idNumar+"\n"+"vizibil_W_mm "+vizibil_W_mm+"\n"+"vizibil_H_mm "+vizibil_H_mm+
    // "\n"+"total_W_mm "+total_W_mm+"\n"+"total_H_mm "+total_H_mm)
    var m_left = myPage.marginPreferences.left; //mmyX2
    var m_right = myPage.marginPreferences.right; //mmyX1
    var m_top = myPage.marginPreferences.top; //mmyy2
    var m_bottom = myPage.marginPreferences.bottom; //mmyy1
    var w = w_t - (2 * m_left); // w vizibil
    var h = h_t - (2 * m_top); // h vizibil

    ///    HW     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    var HW_h_final = (h * 0.1); // 10% hw

    var SAIZECI_Hwai = (HW_h_final * 0.6) // 60% din inaltime lui h

    var Y_HW_final = (m_top + h - (HW_h_final));
    var Y_HW_final_latotal = (h - HW_h_final) + ((h_t - h) / 2);


    var myHW = app.documents[0].rectangles.itemByName("hwul");// HW
    var myHWai = app.documents[0].rectangles.itemByName("hwul").rectangles.itemByName("hwaiul"); // ai ul din HW!!!

    myHW.redefineScaling([1, 1]);
    myHWai.redefineScaling([1, 1]);


    myHW.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left];//succes
    app.documents[0].align(myHWai, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
    myHW.fit(FitOptions.CENTER_CONTENT);



    var myBoundsw = myHW.geometricBounds;  // HW dreptunghi alb
    var myYw = myBoundsw[0];
    var myXw = myBoundsw[1];
    var myHw = myBoundsw[2];
    var myWw = myBoundsw[3];
    var W_hw = myHW.geometricBounds[3] - myHW.geometricBounds[1];
    var H_hw = myHW.geometricBounds[2] - myHW.geometricBounds[0];

    var myBoundsw = myHWai.geometricBounds;  // HW dreptunghi alb
    var myYwai = myBoundsw[0];
    var myXwai = myBoundsw[1];
    var myHwai = myBoundsw[2];
    var myWwai = myBoundsw[3];
    var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
    var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];
    var obj = myHWai
    var anchor = AnchorPoint.CENTER_ANCHOR;
    var ow = obj.geometricBounds[3] - obj.geometricBounds[1];
    var oh = obj.geometricBounds[2] - obj.geometricBounds[0];

    // var anchor = AnchorPoint.CENTER_ANCHOR;
    // var ow = obj.geometricBounds[3] - obj.geometricBounds[1]
    // var oh = obj.geometricBounds[2] - obj.geometricBounds[0];



    if (w > h) { // landscape //H_hwai>h-(h*0.1)
        // alert("landscape")
        var pw = HW_h_final * 0.7 / H_hwai;
        var ph = pw;
        var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: ph });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
        // myHWai.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left];//succes

    }
    if (w < h) { // portrait //W_hwai>w-(w*0.1)
        // alert("portret")
        myHW.redefineScaling([1, 1]);
        myHWai.redefineScaling([1, 1]);
        var pw = W_hw * 0.8 / W_hwai;
        var ph = pw;
        var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: ph });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
    }


    var myBoundsw = myHWai.geometricBounds;  // HW dreptunghi alb
    var myYwai = myBoundsw[0];
    var myXwai = myBoundsw[1];
    var myHwai = myBoundsw[2];
    var myWwai = myBoundsw[3];
    var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
    var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];


    if (W_hwai > w - (w * 0.1)) { // portrait //W_hwai>w-(w*0.1)
        // alert("ratia pe langa 1")
        myHW.redefineScaling([1, 1]);
        myHWai.redefineScaling([1, 1]);
        var pw = W_hw * 0.8 / W_hwai;
        var ph = pw;
        var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: ph });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
    }
    myHWai.fit(FitOptions.CENTER_CONTENT);
    // myHWai.fit(FitOptions.PROPORTIONALLY);
    myHW.fit(FitOptions.CENTER_CONTENT);
    // myHW.fit(FitOptions.PROPORTIONALLY);
    myHW.geometricBounds = [Y_HW_final_latotal, -14.17322835, h_t + 14.17322835, w_t + 14.17322835];



    ///    HW    end  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



    ///////  ----------------------------------------------  IQOS

    ///////-----------dt
    //myDT


    var Hdt_final = ((h - HW_h_final) * 0.04); // 4% dt din h minus hw
    var Y_DT_final = (m_top + h - (HW_h_final + Hdt_final));
    var Y_DT_final_latotal = (h - Hdt_final) + ((h_t - h) / 2);

    var myDT = app.documents[0].rectangles.itemByName("dt4");// HW
    var myDTai = app.documents[0].rectangles.itemByName("dt4").rectangles.itemByName("dt4ai"); // 

    //brand35
    var mybrand35 = app.documents[0].rectangles.itemByName("brand35");
    //_gradient
    var my_gradient = app.documents[0].rectangles.itemByName("_gradient"); // 
    // linedthw
    var mylinedthw = app.documents[0].graphicLines.itemByName("linedthw"); // 


    myDT.redefineScaling([1, 1]);
    myDTai.redefineScaling([1, 1]);
    mylinedthw.redefineScaling([1, 1]);
    mybrand35.redefineScaling([1, 1]);
    my_gradient.redefineScaling([1, 1]);




    var myBoundsDT = myDT.geometricBounds;  // DT dreptunghi alb
    var myYdt = myBoundsDT[0];
    var myXdt = myBoundsDT[1];
    var myHdt = myBoundsDT[2];
    var myWdt = myBoundsDT[3];
    var W_dt = myDT.geometricBounds[3] - myDT.geometricBounds[1];
    var H_dt = myDT.geometricBounds[2] - myDT.geometricBounds[0];

    var myBoundsDTai = myDTai.geometricBounds;  // DTai dreptunghi alb
    var myYdtai = myBoundsDTai[0];
    var myXdtai = myBoundsDTai[1];
    var myHdtai = myBoundsDTai[2];
    var myWdtai = myBoundsDTai[3];
    var W_dtai = myDTai.geometricBounds[3] - myDTai.geometricBounds[1];
    var H_dtai = myDTai.geometricBounds[2] - myDTai.geometricBounds[0];
    var objdtai = myDTai

    myDT.geometricBounds = [Y_DT_final, -14.174, Y_DT_final + (Hdt_final), w_t + 14.174];//
    mylinedthw.geometricBounds = [Y_DT_final + Hdt_final, -14.174, Y_DT_final + Hdt_final, w_t + 14.174];
    mylinedthw.strokeWeight = 1;
    myDT.fit(FitOptions.CENTER_CONTENT);



    ///////-----------dt end



    //brand35 !! portrait
    var myBoundsbrand35 = mybrand35.geometricBounds;  // HW dreptunghi alb
    var myYbrand35 = myBoundsbrand35[0];
    var myXbrand35 = myBoundsbrand35[1];
    var myHbrand35 = myBoundsbrand35[2];
    var myWbrand35 = myBoundsbrand35[3];
    var W_hbrand35 = mybrand35.geometricBounds[3] - mybrand35.geometricBounds[1];
    var H_hbrand35 = mybrand35.geometricBounds[2] - mybrand35.geometricBounds[0];



    //_gradient !! portrait
    var myBounds_gradient = my_gradient.geometricBounds;  // HW dreptunghi alb
    var myYmy_gradient = myBounds_gradient[0];
    var myXmy_gradient = myBounds_gradient[1];
    var myHmy_gradient = myBounds_gradient[2];
    var myWmy_gradient = myBounds_gradient[3];
    // alert(" 1 " + myYmy_gradient + " 2 " + myXmy_gradient + " 3 " + myHmy_gradient + " 4 " + myWmy_gradient)
    var Wmy_gradient = myWmy_gradient - myXmy_gradient;
    var Hmy_gradient = myHmy_gradient - myYmy_gradient;



    var myKeylogo = app.documents[0].polygons.itemByName("_guide50dinlogo_1");
    //myKeylogo !! portrait
    var myBoundsmyKeylogo = myKeylogo.geometricBounds;  // 
    var myYmyKeylogo = myBoundsmyKeylogo[0];
    var myXmyKeylogo = myBoundsmyKeylogo[1];
    var myHmyKeylogo = myBoundsmyKeylogo[2];
    var myWmyKeylogo = myBoundsmyKeylogo[3];
    var W_hmyKeylogo = myKeylogo.geometricBounds[3] - myKeylogo.geometricBounds[1];
    var H_hmyKeylogo = myKeylogo.geometricBounds[2] - myKeylogo.geometricBounds[0];

    myKeylogo.redefineScaling([1, 1]);
    //logoIQOS
    var mylogo = app.documents[0].groups.itemByName("logoIQOS");
    var myBoundsmylogo = mylogo.geometricBounds;  // 
    var myYmylogo = myBoundsmylogo[0];
    var myXmylogo = myBoundsmylogo[1];
    var myHmylogo = myBoundsmylogo[2];
    var myWmylogo = myBoundsmylogo[3];
    var W_hmylogo = mylogo.geometricBounds[3] - mylogo.geometricBounds[1];
    var H_hmylogo = mylogo.geometricBounds[2] - mylogo.geometricBounds[0];
    mylogo.redefineScaling([1, 1]);


    //_guide50dinlogo_3
    var my_guide2 = app.documents[0].polygons.itemByName("_guide50dinlogo_2"); // 


    //_guide50dinlogo_3
    var my_guide3 = app.documents[0].polygons.itemByName("_guide50dinlogo_3"); // 

    //_guide50dinlogo_4
    var my_guide4 = app.documents[0].polygons.itemByName("_guide50dinlogo_4"); // 


    //_guide50dinlogo_8
    var my_guide8 = app.documents[0].polygons.itemByName("_guide50dinlogo_8");
    var myBoundsmy_guide8 = my_guide8.geometricBounds;  // 
    var myYmy_guide8 = myBoundsmy_guide8[0];

    //_guide18%IQOSkey_2
    var my_guide18 = app.documents[0].polygons.itemByName("_guide18%IQOSkey_2");

    var myBoundsmy_guide18 = my_guide18.geometricBounds;  // 
    var myYmy_guide18 = myBoundsmy_guide18[0];
    var myXmy_guide18 = myBoundsmy_guide18[1];
    var myHmy_guide18 = myBoundsmy_guide18[2];
    var myWmy_guide18 = myBoundsmy_guide18[3];
    var Wmy_guide18 = myWmy_guide18 - myXmy_guide18;
    var Hmy_guide18 = myHmy_guide18 - myYmy_guide18;



    //_Disclaimer
    var my_Disclaimer = app.documents[0].textFrames.itemByName("_Disclaimer");
    var myBoundsmy_Disclaimer = my_Disclaimer.geometricBounds;  // 
    var myY_Disclaimer = myBoundsmy_Disclaimer[0];
    var myX_Disclaimer = myBoundsmy_Disclaimer[1];
    var myH_Disclaimer = myBoundsmy_Disclaimer[2];
    var myW_Disclaimer = myBoundsmy_Disclaimer[3];
    var H_Disclaimer = myH_Disclaimer - myY_Disclaimer; // 2-0
    my_Disclaimer.redefineScaling([1, 1]);




    //// aranjare


    // guide 4
    var x_logoMove = w_t - m_right - W_hmylogo - H_hmyKeylogo;
    var y_logoMove = h_t - m_bottom - ((h * 0.1) + ((h - h * 0.1) * 0.04) + H_hmyKeylogo + H_hmylogo);

    var x_my_guide4 = w_t - m_right - H_hmyKeylogo;
    var y_my_guide4 = y_logoMove + H_hmylogo - H_hmylogo / 9;

    var x_my_guide2 = w_t - m_right - W_hmylogo - W_hmyKeylogo + W_hmylogo / 2;
    var y_my_guide2 = h_t - (m_bottom + (h * 0.1) + ((h - h * 0.1) * 0.04) + H_hmyKeylogo); //61.18//(h - h * 0.1) * 0.04

    var x_my_guide3 = w_t - m_right - H_hmyKeylogo;
    var y_my_guide3 = y_logoMove + H_hmylogo - H_hmylogo / 9; //61.18//(h - h * 0.1) * 0.04






    if (h_t > w_t) { // portrait!!






        mybrand35.geometricBounds = [myYbrand35, -14.174, h_t + 14.174, w_t + 14.174];
        //brand35 !! 2
        var myBoundsbrand35_2 = mybrand35.geometricBounds;  // HW dreptunghi alb
        var myYbrand35_2 = myBoundsbrand35_2[0];
        var myXbrand35_2 = myBoundsbrand35_2[1];
        var myHbrand35_2 = myBoundsbrand35_2[2];
        var myWbrand35_2 = myBoundsbrand35_2[3];
        var W_hbrand35_2 = myWbrand35_2 - myXbrand35_2; // 3 1
        var H_hbrand35_2 = myHbrand35_2 - myYbrand35_2; // 2 0


        var x_Disclaimer_portrait = m_left + H_hmyKeylogo;
        var y_Disclaimer_portrait = h_t - H_hbrand35_2 + 14.174 - Hmy_guide18 - H_Disclaimer;
        my_Disclaimer.move([x_Disclaimer_portrait, y_Disclaimer_portrait])



        mylogo.move([x_logoMove, y_logoMove]) //x,y in points

        my_guide2.move([x_my_guide2, y_my_guide2]) //x,y in points
        my_guide3.move([x_my_guide3, y_my_guide3]) //x,y in points


        var myYmy_gradient_portrait = myYmy_gradient;
        var myY2my_gradient_portrait = myYbrand35 + 3;
        my_gradient.geometricBounds = [myYmy_gradient_portrait, -14.174, myY2my_gradient_portrait, w_t + 14.174];

    } else { // landscape

        var x_Disclaimer = m_left + H_hmyKeylogo;
        var y_Disclaimer = h_t - m_bottom - (h * 0.1) - Hdt_final - Hmy_guide18 - H_Disclaimer;
        my_Disclaimer.move([x_Disclaimer, y_Disclaimer]) //x,y in points

        mybrand35.geometricBounds = [-14.174, myXbrand35, h_t + 14.174, w_t + 14.174];

        var myYmy_gradient_landscape = h_t - m_bottom - (h * 0.1) - Hdt_final - Hmy_gradient;
        var myY2my_gradient_landscape = h_t - m_bottom - (h * 0.1) - Hdt_final + 3;
        my_gradient.geometricBounds = [myYmy_gradient_landscape, -14.174, myY2my_gradient_landscape, w_t];


        x_logoMove = w_t - m_right - W_hmylogo - H_hmyKeylogo;
        // mylogo.move([x_logoMove, myYmylogo]) //x,y in points
        mylogo.move([x_logoMove, y_logoMove])

        my_guide2.move([x_my_guide2, y_my_guide2]) //x,y in points
        my_guide3.move([x_my_guide3, y_my_guide3]) //x,y in points
        my_guide8.move([x_my_guide3, myYmy_guide8]) //x,y in points


        // myKeylogo - pt landscape
        var x_myKeylogo = myXmyKeylogo;
        var y_myKeylogo = m_top;

        myKeylogo.move([x_myKeylogo, y_myKeylogo])


        try {

            //_guide50dinlogo_4
            var my_guide4 = app.documents[0].polygons.itemByName("_guide50dinlogo_4");
            var myBounds_my_guide4 = my_guide4.geometricBounds;  // 
            var myY_my_guide4 = myBounds_my_guide4[0];
            var myX_my_guide4 = myBounds_my_guide4[1];
            var myH_my_guide4 = myBounds_my_guide4[2];
            var myW_my_guide4 = myBounds_my_guide4[3];
            var W_h_my_guide4 = myKeylogo.geometricBounds[3] - myKeylogo.geometricBounds[1];
            var H_h_my_guide4 = myKeylogo.geometricBounds[2] - myKeylogo.geometricBounds[0];

            var x_my_guide4 = myX_my_guide4;
            var y_my_guide4 = m_top + H_hmyKeylogo;
            my_guide4.move([x_my_guide4, y_my_guide4])

        } catch (e) {
            // alert(e);


        }





    }





    ///>>>>>>>>>>>>>>>>

    myPage = app.documents[0].pages.item(0);
    with (myPage) {
        var m_top = myPage.marginPreferences.properties.top;
        var m_left = myPage.marginPreferences.properties.left;
        var m_right = myPage.marginPreferences.properties.right;
        var m_bottom = myPage.marginPreferences.properties.bottom;
        var b = myPage.bounds;
        var W_ = b[3] - b[1]; // 
        var H_ = b[2] - b[0]; // 

        app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
        guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: (W_ - m_right) });
        guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: m_left }); // ok
        guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: m_top }); // ok
        guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: (H_ - m_bottom) });

    }



}



function aliniereElemente(vizibil_W, vizibil_H, total_W, total_H) {


    ////***********************************************************************       ALINIERE ELEMENTE STANGA DREAPTA, FIT  *********************************************/


    var myFrames = myDocument.rectangles;

    //***-****** aliniere margini = vizibil


    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignRightLaVizibil") {
            // alert ("match"); // you changes here  
            count++

            app.documents[0].align(myFrames[i], AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); // margin bounds e la marginile paginii adica la vizibil!!!

        }
    }
    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignLeftLaVizibil") {
            // alert ("match"); // you changes here  
            count++

            app.documents[0].align(myFrames[i], AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); // margin bounds e la marginile paginii adica la vizibil!!!

        }
    }
    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignUpLaVizibil") {
            // alert ("match"); // you changes here  
            count++
            app.documents[0].align(myFrames[i], AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); // margin bounds e la marginile paginii adica la vizibil!!!

        }
    }
    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignDownLaVizibil") {
            // alert ("match"); // you changes here  
            count++

            app.documents[0].align(myFrames[i], AlignOptions.BOTTOM_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); // margin bounds e la marginile paginii adica la vizibil!!!



        }
    }

    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignRightLaPagina") {

            count++

            app.documents[0].align(myFrames[i], AlignOptions.RIGHT_EDGES, AlignDistributeBounds.PAGE_BOUNDS); // PAGE bounds 



        }
    }


    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignLeftLaPagina") {

            count++

            app.documents[0].align(myFrames[i], AlignOptions.LEFT_EDGES, AlignDistributeBounds.PAGE_BOUNDS);

        }
    }


    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignHorizontalCenter") {
            // alert ("match"); // you changes here  
            count++
            app.documents[0].align(myFrames[i], AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);

        }
    }


    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "alignVerticalCentersMinusHW") {
            // alert ("match"); // you changes here  
            count++

            app.documents[0].align(myFrames[i], AlignOptions.VERTICAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
            var myPage = app.documents[0].pages.item(0);

            var b = myPage.bounds;
            var W_ = b[3] - b[1]; // 
            var H_ = b[2] - b[0]; // 

            var myBoundsFrame = myFrames[i].geometricBounds;  // HW dreptunghi alb
            var myYF = myBoundsFrame[0];
            var myXF = myBoundsFrame[1];
            var myHF = myBoundsFrame[2];
            var myWF = myBoundsFrame[3];
            var W_hF = myWF - myXF;
            var H_hF = myHF - myYF;
            var moveFrame = (H_ - (H_ * 0.1)) / 2 - (H_hF / 2)
            //  (H_-(H_*0.1)-H_hF)
            myFrames[i].move([myXF, moveFrame])

        }
    }

    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "expandToVizibil") {  // to page          
            // alert ("match expand"); // you changes here  
            count++
            var myPage = app.documents[0].pages.item(0);

            var m_top = myPage.marginPreferences.properties.top;
            var m_left = myPage.marginPreferences.properties.left;
            var m_right = myPage.marginPreferences.properties.right;
            var m_bottom = myPage.marginPreferences.properties.bottom;

            var b = myPage.bounds;
            var W_ = b[3] - b[1]; // 
            var H_ = b[2] - b[0]; // 
            // alert("  "+H_/2.83464567 +"  "+ W_)
            myFrames[i].geometricBounds = [m_top, m_left, H_ - m_bottom, W_ - m_right]; //succes


        }
    }

    var count = 0
    for (var i = 0; i < myFrames.length; i++) {
        if (myFrames[i].label == "expandToBleed") {  // to page          
            // alert ("match expand"); // you changes here  

            var myPage = app.documents[0].pages.item(0);
            // var myBleed = myPage.documentPreferences.properties.bleed;
            count++
            var b = myPage.bounds;
            var W_ = b[3] - b[1]; // 
            var H_ = b[2] - b[0]; // 
            // alert("  "+H_/2.83464567 +"  "+ W_)
            myFrames[i].geometricBounds = [-14.174, -14.174, H_ + 14.174, W_ + 14.174]; //succes // 14.174

        }
    }
    ////***********************************************************************  END     ALINIERE ELEMENTE STANGA DREAPTA, FIT  *********************************************/
}

function lockLayers() {
    ///************* lock layers + active layer */

    try {

        app.documents.item(0).layers.itemByName("id").locked ^= 1;
        app.documents.item(0).layers.itemByName("vizibil").locked ^= 1;
        app.documents.item(0).layers.itemByName("HW").locked ^= 1;
        app.documents.item(0).layers.itemByName("Panel").locked ^= 1;
        app.documents.item(0).layers.itemByName("Image").locked ^= 1;
        app.documents.item(0).layers.itemByName("DTI 4%").locked ^= 1;

    } catch (e) {
        // alert(e);


    }

    ///************* lock layers */
}
//******************************************************************************************************************* END!! ------ FUNCTII
//************************************************************************************************************* END ------ script
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
app.documents[0].viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
app.documents[0].viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;

progressWin.close();
app.documents[0] = app.documents[0];
app.documents[0].close();
alert("job done.")