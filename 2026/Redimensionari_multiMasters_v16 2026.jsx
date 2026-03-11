// https://graphicdesign.stackexchange.com/questions/8523/sharing-indesign-links-between-mac-windows


// var gOsIsMac = (File.fs == "Macintosh") ? true : false;

// if (File.fs == "Windows") {
//     newPath = newPath.replace(/([A-Z])(\/\/)/, "/$1/");
// }
// else if (File.fs == "Macintosh") {
//     newPath = "/Volumes/" + newPath;
// }

/* 

Automated resizing v18

-canged the color for vizibil
-add undefined situatien la citire excel



de facut:
- trebuie un check la master sa verifice daca exista hw cu hwai
- de ce pune un 0 la final la dimensiuni
- daca nu am nevoie de vizibil si total sa scrie doar total, sau daca vizibil == total sa scrie doar total

fixuri:
4 iulie - fix align to bottom-hw-dti labels add smartbleed123 etc
25 aprilie
fix myBleeed_ - ia bleedul din top of document, nu mai e 5 mereu
fix final name

x februarie
fix check number
label al hw si dti

(c) Dan Ichimescu constantindan@gmail.com

last update: 1 aprilie 2025

update: ia din excel, ia cu units cm, mm, px, poate compune final name, poate accepta eticheta la orice


How it works:
1. open one master 
2. start script

Condition:
1. excel header must respect this header:
    column 0 idNumber  | column 1 place/material etc  | column 2 WIDTH (VISIBLE)  | column 3 HEIGHT (VISIBLE)  | column 4 WIDTH (TOTAL)
    | column 5 HEIGHT (TOTAL)  | column 6 measurement units  | column 7 KV name  | column 8 Campain name/short name  | column 9 final name(if)

2. Hw and dti must be copied from example files

*/

/*
de pus sa faca bleed la rectangle sau si bleed la graphicul pe un layer cu force bleed, alt tab in scriptul cu label// force bleed on left right top bottom
de pus align to key object object1 object2 to algin
!! de pus etichete duble!!! ex de aliniat dreapta si jos - este!
de pus forma noua de import excel!!
de facut sa dispara punctele la titlu si parantezele
poate de modificat structura la excel - doar sa adaugi la ce vine
trebuie facut ca in denumire si in info sa apara in mm
eliminate punctele- inlocuite cu -
trebuie sa fac o metoda de verificare tif rapida
*/

// !!info util: Unless you have a reason not to (which you may well have) open the files with visibility set to false doc = app.open(myDoc, false);

if (!app.documents.length > 0) {
  //
  alert("Please open any master file!");
  exit();
}

app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.generalPreferences.pageNumbering = PageNumberingOptions.absolute;

app.scriptPreferences.enableRedraw = true;
app.layoutWindows[0].transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
app.scriptPreferences.userInteractionLevel =
  UserInteractionLevels.NEVER_INTERACT;

var HwPresent, DtiPresent, HwSize, DtiSize, nameOfFoldeResult, substractFromBottom;
var arrErrors = [];

// HwPresent = false; // true;
// DtiPresent = false; //  true;
HwSize = 10 / 100;
DtiSize = 5 / 100;
nameOfFoldeResult = "_result";

var currentDate = new Date();
var dayOfMonth = currentDate.getDate();
var month = currentDate.getMonth() + 1;
var year = currentDate.getFullYear();
var hours = currentDate.getHours();
var minutes = currentDate.getMinutes();
var seconds = currentDate.getSeconds();
var dateAnZiOraMin =
  "_" +
  dayOfMonth +
  "_" +
  month +
  "_" +
  year +
  "_" +
  hours +
  "_" +
  minutes +
  "_" +
  seconds;

var myFile_calea = app.documents[0].filePath;
var myFileName_full = app.documents[0].fullName + "";
var myFileName = app.documents[0].name + "";
// var myFileName_full_length = myFileName_full.length
// var myFileName_length = myFileName.length
var myFileName0 = myFileName.substr(0, myFileName.lastIndexOf("."));
// app.scriptPreferences.enableRedraw = true;
// $.sleep(1000);
app.documents[0].close(SaveOptions.no);
// exit();
// app.scriptPreferences.enableRedraw = false;
var pathFileMaster = myFile_calea; //String(decodeURI(myFile_calea));

var osconnector = "";
// aici windows mac
if (File.fs == "Windows") {
  // pathFileMaster = pathFileMaster.replace(/([A-Z])(\/\/)/, "/$1/");
  osconnector = "\\";
} else if (File.fs == "Macintosh") {
  //   pathFileMaster = "/Volumes/" + pathFileMaster + "/";
  osconnector = "/";
  pathFileMaster = pathFileMaster;
}

// var replacev1 = pathFileMaser.replace(/^\/(.)/, "$1:");
// var replacev0 = replacev1.replace(/\//g, "\\\\");

/// excel

var mask_array = [".xlsx"]; //.xlsx
var excelFilePathArr = find_files(pathFileMaster, mask_array);

// alert(pathFileMaster + "\n" + excelFilePathArr.length);

if (excelFilePathArr.length === 1) {
  excelFilePath = excelFilePathArr[0];
  // modificare mac ###
  //   excelFilePath = decodeURI(
  //     String(excelFilePathArr[0])
  //       .replace(/^\/(.)/, "$1:")
  //       .replace(/\//g, "\\\\")
  //   );
  //   alert(excelFilePath);
} else {
  alert(
    "Please keep one excel file in same folder like master files!\nOr close the opened excel file!"
  );
  exit();
}

// var excelFilePath = File(myFile_calea + "/" + myFileName0 + ".xlsx"); // NU MERRGE
// var excelFilePath = replacev0 + "\\\\" + myFileName0 + ".xlsx"; // old
// alert(excelFilePath);

var LiniaExcel = 0;
// [Optional] the character to use for splitting the columns in the spreadsheed: e.g. semicolon (;) or tab (\t)
// If it isn't set, semicolon will be used by default
var splitChar = ";";
// [Optional] the worksheet number: either string or number. If it isn't set, the first worksheet will be used by default
var sheetNumber = "1";
// var dataList = GetDataFromExcelPC_INDD(excelFilePath, splitChar, sheetNumber); // returns array
var dataList = GetDataFromExcelPC_INDD(excelFilePath, sheetNumber);


checkMyData(dataList);
function checkMyData(arr) {
  // 2-5 numere
  // 6 units
  // alert("The value from Excel is : " + arr[0][2]);

  function isNumber(value) {
    return typeof value === "number" && !isNaN(value);
  }
  for (var i = 0; i < arr.length; i++) {
    // Iterate through the inner array
    for (var j = 2; j < 6; j++) {
      // Check if the element is undefined
      // alert("arr[i][j] " + " i " + i + " j " + j + "  " + arr[i][j])
      // {
      //     if (!Number(arr[i][j])) {
      //         var excelErr = "A dimension is not a number at line " + (i + 2) + " column " + (j + 1);
      //         arrErrors.push(excelErr)
      //         // alert("A dimension is not a number at line " + (i + 2) + " column " + (j + 1))
      //         // exit()

      //     }
      // }

      // Example usage in InDesign script
      var userValueExcel = arr[i][j]; // prompt("Enter a number:", "0");

      if (userValueExcel === undefined) {
        var excelErr =
          "A dimension is not defined at line " +
          (i + 2) +
          " column " +
          (j + 1);
        arrErrors.push(excelErr);
      } else {
        var userValue = userValueExcel.replace(",", ".");
        // alert("The value is : " + userValue);

        // if (isNumber(Number(userValue))) {//
        if (!isNaN(parseFloat(userValue)) && isFinite(userValue)) {
          // alert("The value is a valid number: " + userValue);
          // You can proceed with your numerical operations here
        } else {
          var excelErr =
            "A dimension is not a number at line " +
            (i + 2) +
            " column " +
            (j + 1);
          arrErrors.push(excelErr);
          // alert("The value is not a valid number: " + userValue);
          // Handle the error case here
        }
      }
    }
  }
}
if (arrErrors.length > 0) {
  // alert()

  writeReportsErrors(arrErrors, pathFileMaster);
  exit();
}

var numberOfFiles = dataList.length;
// alert(numberOfFiles)

var arrMasterName1 = new Array();

for (var i = 0; i < dataList.length; i++) {
  var masterName = dataList[i][7];

  var pattern = /.indd/;

  if (pattern.test(masterName)) {
    // console.log("Substring found!");
    arrMasterName1.push(masterName);
  }
}

//**********************************  start succes remove duplicates

var arrMasterName = removeDuplicates(arrMasterName1);
// check if master file exist on hard
// Function to check if a file exists
function fileExists(filePath) {
  var file = new File(filePath);
  return file.exists;
}

for (var i = 0; i < arrMasterName.length; i++) {
  // alert(arrMasterName[i] + "\n" + myFile_calea)
  var filePath = pathFileMaster + "/" + arrMasterName[i];
  if (fileExists(filePath)) {
    // alert("The file exists: " + filePath);
    app.open(File(pathFileMaster + "/" + arrMasterName[i]), false);
  } else {
    alert("The master file does not exist: " + filePath);
    exit();
  }

  // if (File(myFile_calea + "/" + arrMasterName[i]).exists == true) {
  //     app.open(File(myFile_calea + "/" + arrMasterName[i]), false); // false inseamna in background
  // } else {
  //     alert("A master file is not present! check excel name and file on disk")
  // }
  ordoneazaPagini_dupa_Ratie();
  app.documents[0].save();
  app.documents[0].close(SaveOptions.no);
}

// exit();

//********************************** START PROGRESS BAR
var progressWin = CreateProgressBar();
progressWin.show();
progressWin.update(); // poate merge pe windows
progressWin.pb.minvalue = 0;
progressWin.pb.maxvalue = Number(numberOfFiles);
app.scriptPreferences.userInteractionLevel =
  UserInteractionLevels.NEVER_INTERACT;
CreateProgressBar();
function CreateProgressBar() {
  var w = new Window("window", "Redimensionare");
  /*
    w.bounds = [100, 200, 300, 400];
    w.bounds = {left: 100, top: 200, right: 300, bottom: 400}
    w.bounds = "left: 100, top: 200, right: 300, bottom: 400";
    */
  //w.add ("progressbar", undefined, start, stop);
  w.pb = w.add("progressbar", [12, 12, 1250, 25], 0, undefined); //1250
  w.st = w.add("statictext");
  w.st.bounds = [0, 0, 1200, 20];
  w.st.alignment = "left";
  return w;
}
//********************************** END PROGRESS BAR 1

for (var r = 0; r < dataList.length; r++) {
  // r=0 ca am facut la import sa ignore headeru!!

  //**********************************   PROGRESS BAR
  var KVname = dataList[r][7].replace(".indd", ""); //finalFileNameRaw
  var NameOfMyCampain = dataList[r][8];
  var material_places = cleanMydata(dataList[r][1]);
  if (material_places.length <= 2)
    material_places = "";

  var CountNumber = padWithZeros(r + 1, numberOfFiles); // AM PUS +1 CA R E 0 ACU
  // FINAL NAME
  // var finalFileName = cleanMydata(finalFileNameRaw);
  // var finalFileName = CountNumber + "_" + NameOfMyCampain + "_" + material_places + "__" + dataList[r][4] * 10 + "x" + dataList[r][5] * 10 + "_" + dataList[r][2] * 10 + "x" + dataList[r][3] * 10 + "_" + KVname;

  // ok:
  // var finalFileName = CountNumber + "_" + NameOfMyCampain + "_" + material_places + "__"
  // + dataList[r][4] + "x" + dataList[r][5] + "_" + dataList[r][2] + "x" + dataList[r][3] + "_" + KVname;

  // *10 cand ai milimetri!!!
  var unitsofdimensions = dataList[r][6];
  // var addforname = -1;
  // switch (unitsofdimensions) {
  //   case "mm":
  //     addforname = 1; //1 mm =2.83465 point
  //     break;
  //   case "cm":
  //     addforname = 10; //28.3465 //1 mm =2.83465 point
  //     break;
  //   case "px":
  //     addforname = 1; // 1// 0.74999943307122 //1; //1 mm =2.83465 point // 1 px = 0.74999943307122 pt
  //     break;
  //   case "inch":
  //     addforname = 1; //1;
  //     break;
  //   default:
  //     addforname = 1; //1 mm =2.83465 point

  // }
  var dimensionsname = "", totalname = "";

  if (Number(dataList[r][4]) == Number(dataList[r][2]) && Number(dataList[r][5]) == Number(dataList[r][3])) {
    dimensionsname = Number(dataList[r][4]) + "x" + Number(dataList[r][5]);
  } else {
    dimensionsname = Number(dataList[r][4]) + "x" + Number(dataList[r][5]) + "_" + Number(dataList[r][2]) + "x" + Number(dataList[r][3])
  }



  var finalFileName =
    CountNumber +
    "_" +
    NameOfMyCampain +
    "_" +
    material_places +
    "__" +
    dimensionsname +
    "_" + unitsofdimensions +
    "_KV_" +
    KVname;


  // var finalFileName =
  // CountNumber +
  // "_" +
  // NameOfMyCampain +
  // "_" +
  // material_places +
  // "__" +
  // Number(dataList[r][4]) +
  // "x" +
  // Number(dataList[r][5]) +
  // "_" +
  // Number(dataList[r][2]) +
  // "x" +
  // Number(dataList[r][3]) +
  // "_" + unitsofdimensions + "_KV_" +
  // KVname;



  var varInfoCol4 = dataList[r][6]; // units
  var masterNameFile = dataList[r][7];
  // var finalFileName = readedLine[7];
  var idNumar = dataList[r][0];
  // alert("r:" + r + "\n" + "finalFileName:" + "\n" + finalFileName)
  // alert(finalFileName + "\n" + varInfoCol4 + "\n" + masterNameFile + "\n" + idNumar + "\n" + varInfoCol4)
  progressWin.pb.value = r;
  progressWin.st.text =
    "Processing file: \r" +
    finalFileName +
    "\rMaster: " +
    masterNameFile +
    "\r(" +
    (r + 0) +
    " / " +
    numberOfFiles +
    ")";
  progressWin.update();
  //**********************************  PROGRESS BAR END

  // var masterfiletopen=File(myFile_calea + "/" + masterNameFile)

  //********************************** START OPERATII MASTER

  // alert("master file to open:" + "\n" + File(myFile_calea + "\\" + masterNameFile))
  // app.open(File(pathFileMaster + osconnector + masterNameFile), true); // se deschide cu fereastra pt ca sa mearga fittopageandzoom
  app.open(File(pathFileMaster + osconnector + masterNameFile), false); // 

  //**********************************  RATIA MASTER  */

  var myRatia = new Array();
  for (myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
    myPage = app.documents[0].pages.item(myCounter);
    var b = myPage.bounds;
    var W_ = b[3] - b[1]; //
    var H_ = b[2] - b[0]; //
    var ratia = W_ / H_;
    myRatia.push(ratia);
  }

  //**********************************  RATIA MASTER END */

  //**********************************  operatii pe master */
  var myDocument = app.documents[0];
  var myPage_length = app.documents[0].pages.length;
  var myPages = app.documents[0].pages;

  app.documents[0].layers.everyItem().locked = false;
  // app.documents[0].guides.everyItem().remove();

  ///*************** TEST LAYERE VIZIBIL ID SI HW ************************************************************/
  try {
    // app.documents[0].layers.item('Layer 1').name = 'art';
    app.documents[0].layers
      .add({ name: "vizibil", layerColor: UIColors.RED })
      .move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers
      .add({ name: "id", layerColor: UIColors.RED })
      .move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers
      .itemByName("vizibil")
      .move(LocationOptions.AT_BEGINNING);
    app.documents[0].layers.itemByName("id").move(LocationOptions.AT_BEGINNING);
  } catch (e) {
    // alert(e);
  }
  //********************************** TEST LAYERE VIZIBIL ID SI HW  END

  //**********************************  CITESTE W SI H

  var unitsxlx = dataList[r][6]; // units!!
  var changeUnitsOfDocument;

  // Switch statement to handle different export types
  switch (unitsxlx) {
    case "mm":
      // Code for exporting to PDF
      // alert("mm");
      // units = 2.83464567;
      units = 2.83464567; //1 mm =2.83465 point
      changeUnitsOfDocument = "mm";
      // app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      break;

    case "cm":
      //CENTIMETERS
      // alert("cm");
      // units = 28.3464567; //28.3465
      units = 28.3464567; //28.3465 //1 mm =2.83465 point
      changeUnitsOfDocument = "cm";
      // app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.CENTIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.CENTIMETERS;
      break;

    case "px":
      // Code for exporting to JPEG
      // alert("px");
      // units = 0.75;
      units = 1; // 1// 0.74999943307122 //1; //1 mm =2.83465 point // 1 px = 0.74999943307122 pt
      changeUnitsOfDocument = "px";
      // app.scriptPreferences.measurementUnit = MeasurementUnits.PIXELS;
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.PIXELS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.PIXELS;
      break;
    // MeasurementUnits.INCHES //POINTS
    case "inch":
      units = 72; //1;
      changeUnitsOfDocument = "inch";
      // app.scriptPreferences.measurementUnit = MeasurementUnits.INCHES;
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.INCHES;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.INCHES;
      break;

    default:
      // Code for handling other export types
      // alert("mm default");
      units = 2.83464567; //1 mm =2.83465 point
      // app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
  }

  var primu_W_mm = dataList[r][2];
  var primu_W_mm = primu_W_mm.replace(/\,/g, ".");
  var primuH_mm = dataList[r][3];
  var primuH_mm = primuH_mm.replace(/\,/g, ".");
  var aldoilea_W_mm = dataList[r][4];
  var aldoilea_W_mm = aldoilea_W_mm.replace(/\,/g, ".");
  var aldoilea_H_mm = dataList[r][5];
  var aldoilea_H_mm = aldoilea_H_mm.replace(/\,/g, ".");

  var varAria_viz = primu_W_mm * primuH_mm;
  var varAria_total = aldoilea_W_mm * aldoilea_H_mm;

  if (varAria_viz < varAria_total) {
    var vizibil_W = Number(primu_W_mm * units); //* 2.83464567); //mm
    var vizibil_H = Number(primuH_mm * units);
    var total_W = Number(aldoilea_W_mm * units);
    var total_H = Number(aldoilea_H_mm * units);
    var vizibil_W_mm = dataList[r][2];
    var vizibil_H_mm = dataList[r][3];
    var total_W_mm = dataList[r][4];
    var total_H_mm = dataList[r][5];
  } else {
    var vizibil_W = Number(aldoilea_W_mm * units);
    var vizibil_H = Number(aldoilea_H_mm * units);
    var total_W = Number(primu_W_mm * units);
    var total_H = Number(primuH_mm * units);
    var vizibil_W_mm = dataList[r][4];
    var vizibil_H_mm = dataList[r][5];
    var total_W_mm = dataList[r][2];
    var total_H_mm = dataList[r][3];
  }

  //**********************************   CITESTE W SI H END

  var ratia_final = vizibil_W / vizibil_H;

  //********************************** RATIA FINALA PT UN DOCUMENT DECLINAT DIN MASTER

  var myDocument = app.documents[0];
  var myPage_length = app.documents[0].pages.length;
  var myPages = app.documents[0].pages;

  // alert("pagDeExtras: " + pagDeExtras)
  var ratiaMaster = new Array();
  ratiaMaster = myRatia;

  var closestElement = findClosestElement(myRatia, ratia_final); // !! myRatia este ratia din master

  var pagDeExtras = closestElement;

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

  //********************************** cauta ratia end */

  //********************************** extract page */

  for (var i = myPage_length - 1; i >= 0; i--) {
    if (i > pagDeExtras) {
      myPages[i].remove();
    }
    if (i < pagDeExtras) {
      myPages[i].remove();
    }
  }

  var numarRatieRotunjit_u = Number(myRatia[pagDeExtras]);
  var numarRatieRotunjit = numarRatieRotunjit_u.toFixed(3);
  var numeRatieFolder1 = numarRatieRotunjit.toString();
  var numeRatieFolder = numeRatieFolder1.replace(/\./g, "_");
  // alert("numeRatieFolder " + numeRatieFolder)
  // var myFile_calea = app.documents[0].filePath;
  // alert("calea e "+myFile_calea)
  var myFileName_full = app.documents[0].fullName + "";
  var myFileName = app.documents[0].name + "";

  // var myFileName_full_length = myFileName_full.length
  // var myFileName_length = myFileName.length

  var myPage = app.documents[0].pages.item(0);

  try {
    var f = new Folder(pathFileMaster + "/" + nameOfFoldeResult);
    f.create();
  } catch (e) {
    // alert(e);
  }

  //**********************************   MAI JOS SE FAC OPERATII PE PAGINA */

  // operatii
  scaleDocument(vizibil_W, vizibil_H, total_W, total_H);

  resizeDocument(total_W, total_H);

  idsimargineDocument(
    vizibil_W,
    vizibil_H,
    total_W,
    total_H,
    idNumar,
    varInfoCol4
  );

  // if (HwPresent) {
  //   aliniereHwDocument(vizibil_W, vizibil_H, total_W, total_H);
  //   // aliniereHwDocument(); // hw si dti
  // }
  addMyGuides()

  aliniereElemente(vizibil_W, vizibil_H, total_W, total_H);



  lockLayers();

  //**********************************  MAI SUS SE FAC OPERATII PE PAGINA */
  // MeasurementUnits.CENTIMETERS
  // MeasurementUnits.INCHES

  switch (changeUnitsOfDocument) {
    case "px":
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.PIXELS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.PIXELS;
      break;
    case "mm":
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      break;
    case "cm":
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.CENTIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.CENTIMETERS;
      break;
    case "inch":
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.INCHES;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.INCHES;
      break;

    default:
      app.documents[0].viewPreferences.horizontalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
      app.documents[0].viewPreferences.verticalMeasurementUnits =
        MeasurementUnits.MILLIMETERS;
  }

  var savedocname = finalFileName + "_.indd";

  // app.activeWindow.zoom(ZoomOptions.fitPage); // eroare
  // app.layoutWindows[0].zoom(ZoomOptions.fitPage);

  ////------------------ SET SCREEN MODE !! // merge doar cu open cu fereastra normal , cu true nu cu false
  // fittopageandzoom();
  function fittopageandzoom() {
    var doc = app.documents[0];
    app.selection = null;
    var layoutWindow = doc.layoutWindows[0];

    // Fit the page in the window
    layoutWindow.zoom(ZoomOptions.FIT_PAGE);

    // Alternative zoom options:
    // layoutWindow.zoom(ZoomOptions.FIT_SPREAD); // To fit the spread in the window
    // layoutWindow.zoom(ZoomOptions.ACTUAL_SIZE); // To set the zoom to actual size (100%)
    // layoutWindow.zoom(ZoomOptions.FIT_SELECTION); // To fit the selected object in the window

    // Get the current zoom percentage
    var currentZoom = layoutWindow.zoomPercentage;

    // Define the zoom out factor (e.g., 90% of the current zoom level)
    var zoomOutFactor = 0.9;

    // Calculate the new zoom percentage
    var newZoom = currentZoom * zoomOutFactor;

    // Set the new zoom percentage
    layoutWindow.zoomPercentage = newZoom;

    // Set the screen mode to Preview
    layoutWindow.screenMode = ScreenModeOptions.PREVIEW_OFF; //PREVIEW_OFF;
    layoutWindow.screenMode = ScreenModeOptions.PREVIEW_TO_PAGE;
  }

  ////------------------ SET SCREEN MODE !!

  // alert("r:" + r + "\n" + myDocument.name + "\n" + savedocname)
  var savedoc = myDocument.save(
    File(pathFileMaster + "/" + nameOfFoldeResult + "/" + savedocname)
  );
  // alert("r:" + r + "\n" + myDocument.name + "\n" + savedocname)
  savedoc.close(SaveOptions.no);
  // alert("document close")
} //**********************************  END for excel

//********************************** END ------ script

// app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;


progressWin.close();

alert("Job done!\nFiles are in " + nameOfFoldeResult + " folder");

function cleanMydata(finalFileNameRaw) {
  var finalFileName = finalFileNameRaw.replace(/\./g, "_");
  var finalFileName = finalFileName.replace(/\,/g, "_");
  var finalFileName = finalFileName.replace(/\)/g, "_");
  var finalFileName = finalFileName.replace(/\(/g, "_");
  // var finalFileName = finalFileName.replace(/\:/g, "_");
  // var finalFileName = finalFileName.replace(/\;/g, "_");
  // var finalFileName = finalFileName.replace(/`/g, "_");
  // var finalFileName = finalFileName.replace(/\!/g, "_");
  // var finalFileName = finalFileName.replace(/\?/g, "_");
  // var finalFileName = finalFileName.replace(/\>/g, "_");
  // var finalFileName = finalFileName.replace(/\</g, "_");
  // var finalFileName = finalFileName.replace(/\//g, "_");
  // var finalFileName = finalFileName.replace(/\[/g, "_");
  // var finalFileName = finalFileName.replace(/\\/g, "_");
  // var finalFileName = finalFileName.replace(/\|/g, "_");
  // var finalFileName = finalFileName.replace(/\]/g, "_");
  // var finalFileName = finalFileName.replace(/\{/g, "_");
  // var finalFileName = finalFileName.replace(/\}/g, "_");
  // var finalFileName = finalFileName.replace(/\*/g, "_");
  // var finalFileName = finalFileName.replace(/\^/g, "_");
  // var finalFileName = finalFileName.replace(/\$/g, "_");
  // var finalFileName = finalFileName.replace(/\&/g, "_");
  // var finalFileName = finalFileName.replace(/\"/g, "_");
  // var finalFileName = finalFileName.replace(/\'/g, "_");

  // var finalFileName = finalFileNameRaw.replace(/[\/\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");// fara punct
  var finalFileName = finalFileNameRaw.replace(
    /[\/\.\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g,
    "_"
  );
  return finalFileName;
}

function scaleDocument(vizibil_W, vizibil_H, total_W, total_H) {
  app.documents[0].zeroPoint = [0, 0];
  app.documents[0].viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

  var pages = app.documents[0].pages;
  pages[0].marginPreferences.properties = {
    top: 0,
    left: 0,
    bottom: 0,
    right: 0,
  };
  pages[0].layoutRule = LayoutRuleOptions.SCALE;

  pages[0].resize(
    CoordinateSpaces.SPREAD_COORDINATES,
    AnchorPoint.CENTER_ANCHOR,
    ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
    [vizibil_W, vizibil_H]
  );
}

function resizeDocument(total_W, total_H) {
  pages = app.documents[0].pages;
  pages[0].marginPreferences.properties = {
    top: 0,
    left: 0,
    bottom: 0,
    right: 0,
  };
  pages[0].layoutRule = LayoutRuleOptions.OFF;
  pages[0].resize(
    CoordinateSpaces.SPREAD_COORDINATES,
    AnchorPoint.CENTER_ANCHOR,
    ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
    [total_W, total_H]
  );
}

function idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar) {
  try {
    app.documents[0].activeLayer = app.documents[0].layers.itemByName("id");
    app.documents.item(0).layers.itemByName("id").locked = false;
    app.documents.item(0).layers.itemByName("vizibil").locked = false;
  } catch (e) { }

  myPage = app.documents[0].pages.item(0);
  var textFrame_info = myPage.textFrames.add();
  textFrame_info.geometricBounds = [0, total_W + 50, 300, total_W + 500];

  textFrame_info.contents =
    "W vizibil = " +
    vizibil_W_mm +
    "\n" +
    "H vizibil = " +
    vizibil_H_mm +
    "\n" +
    "W total = " +
    total_W_mm +
    "\n" +
    "H total = " +
    total_H_mm +
    "\n" +
    " - " +
    varInfoCol4 +
    "\n" +
    " - " +
    masterNameFile;

  textFrame_info.label = "Info";
  var myText = textFrame_info.parentStory.paragraphs.item(0);
  myText.paragraphs.everyItem().properties = {
    appliedFont: app.fonts.item("Times"), //Times //IQOS Sans
    fontStyle: "Regular",
    pointSize: 25,
    leading: 30,
    fillColor: "Black",
  };

  try {
    appliedLanguage = "English: USA";
    textFrame_info.fit(FitOptions.FRAME_TO_CONTENT);
    textFrame_info.move([total_W + 20, 0]); //total_H + 15
  } catch (e) { }

  //********************************** text info end

  var m_top = (total_H - vizibil_H) / 2;
  var m_left = (total_W - vizibil_W) / 2;
  var m_right = m_left;
  var m_bottom = m_top;

  var myPage = app.documents[0].pages.item(0);
  myPage.marginPreferences.properties = {
    top: m_top,
    left: m_left,
    right: m_right,
    bottom: m_bottom,
  };

  app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
  var myItem_vizibil = myPage.rectangles.add({
    geometricBounds: [m_top, m_left, m_top + vizibil_H, m_left + vizibil_W],
  });

  var myColorforvizibil = app.documents[0].colors.add({
    name: "colorforvisible",
    // colorValue: [0, 0, 100, 0],
    colorValue: [0, 100, 0, 0],
    model: ColorModel.PROCESS,
    space: ColorSpace.CMYK,
  });
  myItem_vizibil.label = "Vizibil";
  myItem_vizibil.name = "RectangleVizibil";
  myItem_vizibil.fillColor = "None";
  myItem_vizibil.strokeColor = myColorforvizibil;
  myItem_vizibil.strokeWeight = 1;
  myItem_vizibil.strokeTint = 100;
  myItem_vizibil.strokeAlignment = StrokeAlignment.INSIDE_ALIGNMENT;
  // myItem_vizibil.overprintFill = false;
  myItem_vizibil.transparencySettings.blendingSettings.opacity = 100;
  myItem_vizibil.transparencySettings.blendingSettings.blendMode =
    BlendMode.NORMAL;
}

// function aliniereHwDocument() {
//   //vizibil_W, vizibil_H, total_W, total_H

//   var myPage = app.documents[0].pages.item(0);
//   var b_pgebounds = myPage.bounds;

//   var w_t = b_pgebounds[3] - b_pgebounds[1];
//   var h_t = b_pgebounds[2] - b_pgebounds[0];

//   var m_left = myPage.marginPreferences.left; //mmyX2
//   var m_right = myPage.marginPreferences.right; //mmyX1
//   var m_top = myPage.marginPreferences.top; //mmyy2
//   var m_bottom = myPage.marginPreferences.bottom; //mmyy1
//   var w = w_t - 2 * m_left; // w vizibil
//   var h = h_t - 2 * m_top; // h vizibil

//   //**********************************    HW     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

//   var HW_h_final = h * HwSize; //0.1); // 10% hw
//   var SAIZECI_Hwai = HW_h_final * 0.6; // 60% din inaltime lui h
//   var Y_HW_final = m_top + h - HW_h_final;
//   var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;
//   var myHW = app.documents[0].rectangles.itemByName("hwul"); // HW
//   var myHWai = app.documents[0].rectangles
//     .itemByName("hwul")
//     .rectangles.itemByName("hwaiul"); // ai ul din HW!!!
//   myHW.redefineScaling([1, 1]);
//   myHWai.redefineScaling([1, 1]);

//   myHW.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left]; //succes
//   app.documents[0].align(
//     myHWai,
//     AlignOptions.HORIZONTAL_CENTERS,
//     AlignDistributeBounds.MARGIN_BOUNDS
//   );
//   myHW.fit(FitOptions.CENTER_CONTENT);

//   var myBoundsw = myHW.geometricBounds; // HW dreptunghi alb
//   var myYw = myBoundsw[0];
//   var myXw = myBoundsw[1];
//   var myHw = myBoundsw[2];
//   var myWw = myBoundsw[3];
//   var W_hw = myHW.geometricBounds[3] - myHW.geometricBounds[1];
//   var H_hw = myHW.geometricBounds[2] - myHW.geometricBounds[0];

//   var myBoundsw = myHWai.geometricBounds; // HW dreptunghi alb
//   var myYwai = myBoundsw[0];
//   var myXwai = myBoundsw[1];
//   var myHwai = myBoundsw[2];
//   var myWwai = myBoundsw[3];
//   var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
//   var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];
//   var obj = myHWai;
//   var anchor = AnchorPoint.CENTER_ANCHOR;
//   var ow = obj.geometricBounds[3] - obj.geometricBounds[1];
//   var oh = obj.geometricBounds[2] - obj.geometricBounds[0];

//   if (w > h) {
//     // landscape

//     var pw = (HW_h_final * 0.7) / H_hwai;
//     var ph = pw;
//     var matrix = app.transformationMatrices.add({
//       horizontalScaleFactor: pw,
//       verticalScaleFactor: ph,
//     });
//     obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
//   }
//   if (w < h) {
//     myHW.redefineScaling([1, 1]);
//     myHWai.redefineScaling([1, 1]);
//     var pw = (W_hw * 0.8) / W_hwai;
//     var ph = pw;
//     var matrix = app.transformationMatrices.add({
//       horizontalScaleFactor: pw,
//       verticalScaleFactor: ph,
//     });
//     obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
//   }

//   var myBoundsw = myHWai.geometricBounds; // HW dreptunghi alb
//   var myYwai = myBoundsw[0];
//   var myXwai = myBoundsw[1];
//   var myHwai = myBoundsw[2];
//   var myWwai = myBoundsw[3];
//   var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
//   var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];

//   if (W_hwai > w - w * 0.1) {
//     // portrait //W_hwai>w-(w*0.1)

//     myHW.redefineScaling([1, 1]);
//     myHWai.redefineScaling([1, 1]);
//     var pw = (W_hw * 0.8) / W_hwai;
//     var ph = pw;
//     var matrix = app.transformationMatrices.add({
//       horizontalScaleFactor: pw,
//       verticalScaleFactor: ph,
//     });
//     obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
//   }
//   myHWai.fit(FitOptions.CENTER_CONTENT);

//   myHW.fit(FitOptions.CENTER_CONTENT);

//   myHW.geometricBounds = [
//     Y_HW_final_latotal,
//     -14.17322835,
//     h_t + 14.17322835,
//     w_t + 14.17322835,
//   ];

//   //**********************************     HW    end  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



// } /// end aliniere hw


function addMyGuides() {
  //**********************************     Guides  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  myPage = app.documents[0].pages.item(0);
  with (myPage) {
    var m_top = myPage.marginPreferences.properties.top;
    var m_left = myPage.marginPreferences.properties.left;
    var m_right = myPage.marginPreferences.properties.right;
    var m_bottom = myPage.marginPreferences.properties.bottom;
    var b = myPage.bounds;
    var W_ = b[3] - b[1]; //
    var H_ = b[2] - b[0]; //

    app.documents[0].activeLayer =
      app.documents[0].layers.itemByName("vizibil");
    guides.add(undefined, {
      orientation: HorizontalOrVertical.vertical,
      location: W_ - m_right,
    });
    guides.add(undefined, {
      orientation: HorizontalOrVertical.vertical,
      location: m_left,
    }); // ok
    guides.add(undefined, {
      orientation: HorizontalOrVertical.horizontal,
      location: m_top,
    }); // ok
    guides.add(undefined, {
      orientation: HorizontalOrVertical.horizontal,
      location: H_ - m_bottom,
    });
  }

  //**********************************     Guides end  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
}


function aliniereElemente(vizibil_W, vizibil_H, total_W, total_H) {
  // + redefine Scaling

  //**********************************       ALINIERE ELEMENTE STANGA DREAPTA, FIT

  var myPage = app.documents[0].pages.item(0);
  var pageItems = myPage.allPageItems;
  var m_top = myPage.marginPreferences.properties.top;
  var m_left = myPage.marginPreferences.properties.left;
  var m_right = myPage.marginPreferences.properties.right;
  var m_bottom = myPage.marginPreferences.properties.bottom;
  var b = myPage.bounds;
  var W_ = b[3] - b[1]; //
  var H_ = b[2] - b[0]; //
  var H_vizibil = H_ - (m_top + m_bottom);
  var W_vizibil = W_ - (m_left + m_right);

  var doc = app.documents[0];
  var items = doc.allPageItems;

  var labelToCheckA = "Health_Warning";  // Change as needed
  var labelToCheckB = "DeTeIi";  // Change as needed

  var foundLabelHW = false;
  var foundLabelDTI = false;
  // var dtipresent;
  var myDTIitem, myHwItem, objBase;
  // Loop through all page items
  for (var i = 0; i < items.length; i++) {
    var label = items[i].label;

    if (!label) continue;

    if (label === labelToCheckA) {
      foundLabelHW = true;
      myHwItem = items[i]
    }
    if (label === labelToCheckB) {
      foundLabelDTI = true;
      myDTIitem = items[i]
    }

    // Optional: stop early if both found
    if (foundLabelHW && foundLabelDTI) break;
  }

  if (foundLabelHW && foundLabelDTI) {
    substractFromBottom = H_vizibil * HwSize + (H_vizibil - (H_vizibil * HwSize)) * DtiSize;
    // objBase = myDTIitem
    //H_vizibil * (HwSize + DtiSize);
    // alert(substractFromBottom);
  } else if (foundLabelHW && !foundLabelDTI) {
    substractFromBottom = H_vizibil * HwSize;
    // objBase = myHwItem
    // alert(substractFromBottom);
  } else {
    substractFromBottom = 0;
  }


  var zeceLaSuta = H_vizibil * HwSize;
  var DeTeiRealSize = (H_vizibil - zeceLaSuta) * DtiSize;
  var douaZeciLaSuta = H_vizibil * 0.2;

  //key 80% din Width  in portrait
  //key 85% din H-10%dinH  in landscape
  var key_optezeci_portret = W_vizibil * 0.8;
  var key_optezecisicinci_landscape = (H_vizibil - (H_vizibil * 0.1)) * 0.85;


  var myObjectList = new Array();


  for (var myCounter = 0; myCounter < pageItems.length; myCounter++) {
    switch (pageItems[myCounter].constructor.name) {
      case "Rectangle":
      case "Oval":
      case "Polygon":
      case "TextFrame":
      case "Group":
      case "Button":
      case "GraphicLine":
        myObjectList.push(pageItems[myCounter]);
        break;
    }
  }
  ///////=============


  // Define your list of label strings and associated actions
  var labelActions = {

    "Health_Warning": function (myObjectToAling) {
      var mychild = myObjectToAling.rectangles.item(0);
      aliniere_zece_la_suta(myObjectToAling, mychild)
    },
    "DeTeIi": function (myObjectToAling) {
      var mychild = myObjectToAling.rectangles.item(0);
      aliniere_cinci_la_suta(myObjectToAling, mychild)
    },
    "linia": function (myObjectToAling) {
      aliniere_linia(myObjectToAling)
    },

    ////>>> - HW si dti

    "alignTo_Visible_Bottom-HW": function (myObjectToAling) {
      // alert("dtipresent " + dtipresent)

      // moveObjectAbove(myObjectToAling, objBase)

      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var newTopY = H_ - (m_bottom + boundsofMyObject[5] + substractFromBottom);
      var currentTopY = boundsofMyObject[0];
      var deltaY = newTopY - currentTopY;
      var deltaX = 0; // Keep original X position

      myObjectToAling.move(undefined, [deltaX, deltaY]);



      // var moverBounds = mover.geometricBounds;
      // var baseBounds = base.geometricBounds;

      // var moverHeight = moverBounds[2] - moverBounds[0];

      // // Calculate new position: just above the base object
      // var newTopY = baseBounds[0] - moverHeight;
      // var currentTopY = moverBounds[0];

      // var deltaY = newTopY - currentTopY;
      // var deltaX = 0; // Keep original X position

      // myObjectToAling.move(undefined, [deltaX, deltaY]);




      // myObjectToAling.move([boundsofMyObject[1], moveFrame2]);


      // myObjectToAling.move([boundsofMyObject[1], moveFrame2]);
    },
    "alignTo_Visible_Corner_leftBottom-HW": function (myObjectToAling) {

      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var newTopY = H_ - (m_bottom + boundsofMyObject[5] + substractFromBottom);
      var currentTopY = boundsofMyObject[0];
      var deltaY = newTopY - currentTopY;
      var deltaX = 0; // Keep original X position

      myObjectToAling.move(undefined, [deltaX, deltaY]);

      // moveObjectAbove(myObjectToAling, objBase)
      // var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      // var moveFrame = H_ - (m_bottom + boundsofMyObject[5] + substractFromBottom);
      // myObjectToAling.move([boundsofMyObject[1], moveFrame]);

      app.documents[0].align(
        myObjectToAling,
        AlignOptions.LEFT_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      ); //
    },
    "alignTo_Visible_Corner_rightBottom-HW": function (myObjectToAling) {
      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var newTopY = H_ - (m_bottom + boundsofMyObject[5] + substractFromBottom);
      var currentTopY = boundsofMyObject[0];
      var deltaY = newTopY - currentTopY;
      var deltaX = 0; // Keep original X position

      myObjectToAling.move(undefined, [deltaX, deltaY]);
      // moveObjectAbove(myObjectToAling, objBase)
      // var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      // var moveFrame = H_ - (m_bottom + boundsofMyObject[5] + substractFromBottom);
      // myObjectToAling.move([boundsofMyObject[1], moveFrame]);

      app.documents[0].align(
        myObjectToAling,
        AlignOptions.RIGHT_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      );
    },



    ////>>> - HW si dti


    "cheia": function (myObjectToAling) {
      {

        var obj = myObjectToAling;
        var anchor = AnchorPoint.CENTER_ANCHOR;

        var H_myObjectToAling = obj.geometricBounds[2] - obj.geometricBounds[0];

        //key 80% din Width  in portrait
        //key 85% din H-10%dinH  in landscape
        // key_optezeci_portret;
        //  key_optezecisicinci_landscape

        if (H_ > W_) { // portret

          var pw = (key_optezeci_portret) / H_myObjectToAling;
          var ph = pw;
          var matrix = app.transformationMatrices.add({
            horizontalScaleFactor: pw,
            verticalScaleFactor: ph,
          });
          obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);

          //AlignOptions.VERTICAL_CENTERS


          app.documents[0].align(
            myObjectToAling,
            AlignOptions.HORIZONTAL_CENTERS,
            AlignDistributeBounds.MARGIN_BOUNDS
          );
          app.documents[0].align(
            myObjectToAling,
            AlignOptions.VERTICAL_CENTERS,
            AlignDistributeBounds.MARGIN_BOUNDS
          );
          // la jumate intre top si hw
          var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
          var moveFrame = ((H_ - m_bottom - (H_vizibil * 0.1)) / 2) - (boundsofMyObject[5] / 2);
          myObjectToAling.move([boundsofMyObject[1], moveFrame]);

        } else {
          var pw = (key_optezecisicinci_landscape) / H_myObjectToAling;
          var ph = pw;
          var matrix = app.transformationMatrices.add({
            horizontalScaleFactor: pw,
            verticalScaleFactor: ph,
          });
          obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);


          app.documents[0].align(
            myObjectToAling,
            AlignOptions.HORIZONTAL_CENTERS,
            AlignDistributeBounds.MARGIN_BOUNDS
          );
          app.documents[0].align(
            myObjectToAling,
            AlignOptions.VERTICAL_CENTERS,
            AlignDistributeBounds.MARGIN_BOUNDS
          );
          // la landscape sta pe linia de 20%
          var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
          var moveFrame = ((H_ - m_bottom - (H_vizibil * 0.2))) - (boundsofMyObject[5]);
          myObjectToAling.move([boundsofMyObject[1], moveFrame]);


        }

        var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
        var lamijloculcheiei = (boundsofMyObject[2] - boundsofMyObject[0]) / 2 + boundsofMyObject[0]


        with (myPage) {


          guides.add(app.documents[0].layers.itemByName("Image"), {
            orientation: HorizontalOrVertical.horizontal,
            location: lamijloculcheiei
          }); // ok
        }


      }
    },


    "douzecilasuta": function (myObjectToAling) {
      {

        var obj = myObjectToAling;
        var anchor = AnchorPoint.CENTER_ANCHOR;

        var H_myObjectToAling = obj.geometricBounds[2] - obj.geometricBounds[0];

        var pw = (douaZeciLaSuta) / H_myObjectToAling;
        var ph = pw;
        var matrix = app.transformationMatrices.add({
          horizontalScaleFactor: pw,
          verticalScaleFactor: ph,
        });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);


        app.documents[0].align(
          myObjectToAling,
          AlignOptions.BOTTOM_EDGES,
          AlignDistributeBounds.MARGIN_BOUNDS
        );

        var boundsofMyObject = getBoundsOfMyObject(obj);
        var locationofMyGuidesFor20zecilasuta = H_ - boundsofMyObject[5] - m_bottom


        with (myPage) {


          guides.add(app.documents[0].layers.itemByName("Image"), {
            orientation: HorizontalOrVertical.horizontal,
            location: locationofMyGuidesFor20zecilasuta
          }); // ok
        }
        // app.activeDocument.guides.add
        //     (
        //         // First argument, the layer:
        //         app.activeDocument.layers.itemByName("Image")

        //         // Do not miss out the comma, that parts the two arguments:
        //         ,

        //         // Second argument, all other properties written as an object:
        //         {
        //             location: "20 mm",
        //             orientation: HorizontalOrVertical.VERTICAL, //VERTICAL, HORIZONTAL

        //             fitToPage: true,
        //             guideColor: UIColors.GOLD,
        //             guideType: GuideTypeOptions.RULER
        //         }
        //     );


      }
    },

    "alignToPageBottom": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.BOTTOM_EDGES,
        AlignDistributeBounds.PAGE_BOUNDS
      );
    },
    "alignToPageLeft": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.LEFT_EDGES,
        AlignDistributeBounds.PAGE_BOUNDS
      );
    },
    "alignToPageRight": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.RIGHT_EDGES,
        AlignDistributeBounds.PAGE_BOUNDS
      ); // PAGE bounds
    },
    "alignToPageTop": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.TOP_EDGES,
        AlignDistributeBounds.PAGE_BOUNDS
      ); // PAGE bounds
    },

    "alignTo_Visible_Corner_leftTop": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.TOP_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      );
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.LEFT_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      );
    },

    "alignTo_Visible_Corner_RightTop": function (myObjectToAling) {
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.TOP_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      );
      app.documents[0].align(
        myObjectToAling,
        AlignOptions.RIGHT_EDGES,
        AlignDistributeBounds.MARGIN_BOUNDS
      );
    },
    "alignTo_Visible_Left": function (myObjectToAling) {
      {
        app.documents[0].align(
          myObjectToAling,
          AlignOptions.LEFT_EDGES,
          AlignDistributeBounds.MARGIN_BOUNDS
        );
      }

    },
    "alignTo_Visible_Right": function (myObjectToAling) {
      {
        app.documents[0].align(
          myObjectToAling,
          AlignOptions.RIGHT_EDGES,
          AlignDistributeBounds.MARGIN_BOUNDS
        );
      }
    },
    "alignTo_Visible_Top": function (myObjectToAling) {
      {
        app.documents[0].align(
          myObjectToAling,
          AlignOptions.TOP_EDGES,
          AlignDistributeBounds.MARGIN_BOUNDS
        );
      }
    },
    "alignVerticalCentersMinusHW": function (myObjectToAling) {
      {
        // alert("alignVerticalCentersMinusHW")
        var moveFrame = (H_ - H_ * HwSize) / 2 - H_obj / 2; //HwSize= 0.1

        myObjectToAling.move([myXF, moveFrame]);
      }
    },
    "alignHorizontalCenter": function (myObjectToAling) {
      {
        app.documents[0].align(
          myObjectToAling,
          AlignOptions.HORIZONTAL_CENTERS,
          AlignDistributeBounds.MARGIN_BOUNDS
        );
      }
    },
    "expandToVizibil": function (myObjectToAling) {
      {
        myObjectToAling.geometricBounds = [
          m_top,
          m_left,
          H_ - m_bottom,
          W_ - m_right,
        ];
      }
    },
    "expandToBleed": function (myObjectToAling) {

      // alert("expandToBleed")
      var myBleeed_ = myDocument.documentPreferences.properties.documentBleedTopOffset;
      myObjectToAling.geometricBounds = [
        -myBleeed_,
        -myBleeed_,
        H_ + myBleeed_,
        W_ + myBleeed_,
      ]; //succes // 14.174

      // myObjectToAling.geometricBounds = [
      //   -14.174,
      //   -14.174,
      //   H_ + 14.174,
      //   W_ + 14.174,
      // ]; //succes // 14.174



      // aici fix bleed
      // fixbleed(myObjectToAling);

    },


    //  1
    // _____
    //4|   |2
    // _____
    //  3

    //-------

    "expandToBleed123": function (myObjectToAling) {

      // 
      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var myBleeed_ = myDocument.documentPreferences.properties.documentBleedTopOffset;
      myObjectToAling.geometricBounds = [
        -myBleeed_,
        boundsofMyObject[1],
        H_ + myBleeed_,
        W_ + myBleeed_,
      ]; //succes // 14.174

      // aici fix bleed
      // fixbleed(myObjectToAling);

    },

    "expandToBleed234": function (myObjectToAling) {

      // 
      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var myBleeed_ = myDocument.documentPreferences.properties.documentBleedTopOffset;
      myObjectToAling.geometricBounds = [
        boundsofMyObject[0],
        -myBleeed_,
        H_ + myBleeed_,
        W_ + myBleeed_,
      ]; //succes // 14.174

      // aici fix bleed
      // fixbleed(myObjectToAling);

    },

    "expandToBleed14": function (myObjectToAling) {

      // alert("expandToBleed")
      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var myBleeed_ = myDocument.documentPreferences.properties.documentBleedTopOffset;
      myObjectToAling.geometricBounds = [
        -myBleeed_,
        -myBleeed_,
        boundsofMyObject[2],
        boundsofMyObject[3],
      ]; //succes // 14.174

      // aici fix bleed
      // fixbleed(myObjectToAling);

    },
    //geometricBounds returns an array of four numbers: [y1, x1, y2, x2]
    "expandToBleed42": function (myObjectToAling) {

      // alert("expandToBleed")
      var boundsofMyObject = getBoundsOfMyObject(myObjectToAling);
      var myBleeed_ = myDocument.documentPreferences.properties.documentBleedTopOffset;
      myObjectToAling.geometricBounds = [
        boundsofMyObject[0],
        -myBleeed_,
        boundsofMyObject[2],
        W_ + myBleeed_,
      ]; //succes // 14.174

      // aici fix bleed
      // fixbleed(myObjectToAling);

    },
    "smartBleed": function (myObjectToAling) {
      // var anchor = app.activeWindow.transformReferencePoint; // test eroare
      var myBleeed_ =
        myDocument.documentPreferences.properties.documentBleedTopOffset;
      var bleedPagina = myBleeed_;
      b_pgebounds = myPage.bounds;
      Wp = b_pgebounds[3] - b_pgebounds[1];
      Hp = b_pgebounds[2] - b_pgebounds[0];
      Wp_mm = (b_pgebounds[3] - b_pgebounds[1]) / 2.8346438836889;
      Hp_mm = (b_pgebounds[2] - b_pgebounds[0]) / 2.8346438836889;
      var mypagY1 = b_pgebounds[0];
      var mypagX1 = b_pgebounds[1];
      var mypagY2 = b_pgebounds[2];
      var mypagX2 = b_pgebounds[3];
      var jumatepaginaW = Wp / 2 - Wp * 0.05;
      var jumatepaginaH = Hp / 2 - Hp * 0.05;
      var zonademijlocpaginaW = Wp * 0.33;
      var zonademijlocpaginaH = Hp * 0.33;
      var Y1H = (Hp - Hp * 0.33) / 2;
      var X1H = 0;
      var Y2H = Hp * 0.33 + Y1H;
      var X2H = Wp;
      var Y1V = 0;
      var X1V = (Wp - Wp * 0.33) / 2;
      var Y2V = Hp;
      var X2V = Wp * 0.33 + X1V;
      // var obj_select = app.selection;
      var obj_select = myObjectToAling;
      // for (i = 0; i < obj_select.length; i++) {
      var obj = obj_select;
      if (obj instanceof Rectangle) {
        var ow = obj.geometricBounds[3] - obj.geometricBounds[1];
        var oh = obj.geometricBounds[2] - obj.geometricBounds[0];
        var myBounds1 = obj.geometricBounds;
        var Y1 = myBounds1[0];
        var X1 = myBounds1[1];
        var Y2 = myBounds1[2];
        var X2 = myBounds1[3];
        obj.fit(FitOptions.frameToContent);
        var myBounds_fit = obj.geometricBounds;
        var Y1f = myBounds_fit[0];
        var X1f = myBounds_fit[1];
        var Y2f = myBounds_fit[2];
        var X2f = myBounds_fit[3];
        var ow_fit = obj.geometricBounds[3] - obj.geometricBounds[1];
        var oh_fit = obj.geometricBounds[2] - obj.geometricBounds[0];
        if (Y1f < -bleedPagina) {
          var caseY1 = -bleedPagina;
        } else {
          if (Y1f > -bleedPagina && Y1f < Hp * 0.05) {
            //
            var dif = Math.abs(Y1f);
            var H_final = Y2f + bleedPagina;
            var ph = H_final / oh_fit;
            var pw = ph;
            // app.layoutWindows[0].transformReferencePoint =
            //     AnchorPoint.BOTTOM_CENTER_ANCHOR;
            // resizeObiect(pw, ph);
            //
            var caseY1 = -bleedPagina;
          } else {
            var caseY1 = Y1;
          }
        }
        if (X1f < -bleedPagina) {
          var caseX1 = -bleedPagina;
        } else {
          if (X1f > -bleedPagina && X1f < Wp * 0.05) {
            //
            var dif = Math.abs(X1f);
            var W_final = X2f + bleedPagina;
            var ph = W_final / ow_fit;
            var pw = ph;
            // app.layoutWindows[0].transformReferencePoint =
            //     AnchorPoint.RIGHT_CENTER_ANCHOR;
            //resizeObiect(pw, ph);
            //
            var caseX1 = -bleedPagina;
          } else {
            var caseX1 = X1;
          }
        }
        if (Y2f > Hp + bleedPagina) {
          var caseY2 = Hp + bleedPagina;
        } else {
          if (Y2f < Hp + bleedPagina && Y2f > Hp - Hp * 0.05) {
            //
            var dif = Hp + bleedPagina - Y2f;
            var H_final = oh_fit + dif;
            var ph = H_final / oh_fit;
            var pw = ph;
            // app.layoutWindows[0].transformReferencePoint =
            //     AnchorPoint.TOP_CENTER_ANCHOR;
            //resizeObiect(pw, ph);
            //
            var caseY2 = Hp + bleedPagina;
          } else {
            var caseY2 = Y2;
          }
        }
        if (X2f > Wp + bleedPagina) {
          var caseX2 = Wp + bleedPagina;
        } else {
          if (X2f < Wp + bleedPagina && X2f > Wp - Wp * 0.05) {
            //
            var dif = Wp + bleedPagina - X2f;
            var W_final = ow_fit + dif;
            var ph = W_final / ow_fit;
            var pw = ph;
            // app.layoutWindows[0].transformReferencePoint =
            //     AnchorPoint.LEFT_CENTER_ANCHOR;
            //resizeObiect(pw, ph);

            var caseX2 = Wp + bleedPagina;
          } else {
            var caseX2 = X2;
          }
        }
        obj.geometricBounds = [caseY1, caseX1, caseY2, caseX2];
        // app.layoutWindows[0].transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
        // function resizeObiect(pw, ph) {
        // }
      }
      // }
      // app.activeDocument.viewPreferences.horizontalMeasurementUnits =
      // app.activeDocument.viewPreferences.verticalMeasurementUnits =
      // MeasurementUnits.millimeters;


    }

  };

  // // Get the active document
  // var doc = app.documents[0];
  /// loop 1
  for (var i = 0; i < myObjectList.length; i++) {
    var myObjectToAling = myObjectList[i];
    // app.select(myObjectToAling)
    // var myObj = app.selection[0];

    // redefinescaling!!
    myObjectToAling.redefineScaling([1, 1]);

    var myBoundsFrame = myObjectList[i].geometricBounds; //
    var myYF = myBoundsFrame[0];
    var myXF = myBoundsFrame[1];
    var myHF = myBoundsFrame[2];
    var myWF = myBoundsFrame[3];
    var W_obj = myWF - myXF;
    var H_obj = myHF - myYF;

    var label = myObjectToAling.label;



    if (label && label.length > 0) {
      var labels = label.split("+");

      for (var j = 0; j < labels.length; j++) {
        // var trimmedLabel = labels[j].trim();
        var trimmedLabel = labels[j].replace(/^\s+|\s+$/g, "");

        if (labelActions.hasOwnProperty(trimmedLabel)) {
          labelActions[trimmedLabel](myObjectToAling);
        }
      }
    }
    // Check if label matches one of the keys
    // else {
    //   if (labelActions.hasOwnProperty(mylabel)) {
    //     try {
    //       labelActions[mylabel](myObjectToAling); // Call associated action
    //     } catch (e) {
    //       //  alert("Error processing item with label '" + label + "': " + e);
    //     }
    //   }
    // }


  }

  // // === Main Script ===
  // function processLabeledItems() {
  //   var doc = app.activeDocument;
  //   var allItems = doc.allPageItems;

  //   for (var i = 0; i < allItems.length; i++) {
  //     var item = allItems[i];
  //     var label = item.label;

  //     if (label && label.length > 0) {
  //       var labels = label.split("+");

  //       for (var j = 0; j < labels.length; j++) {
  //         var trimmedLabel = labels[j].trim();

  //         if (labelActions.hasOwnProperty(trimmedLabel)) {
  //           labelActions[trimmedLabel](item);
  //         }
  //       }
  //     }
  //   }

  //   alert("Processing complete!");
  // }



  // Loop through all page items
  // for (var i = 0; i < doc.allPageItems.length; i++) {
  //   var item = doc.allPageItems[i];
  //   var label = item.label;

  //   // Check if label matches one of the keys
  //   if (labelActions.hasOwnProperty(label)) {
  //     try {
  //       labelActions[label](item); // Call associated action
  //     } catch (e) {
  //       //  alert("Error processing item with label '" + label + "': " + e);
  //     }
  //   }
  // }
  ///////=============


  //**********************************  END     ALINIERE ELEMENTE STANGA DREAPTA, FIT  *********************************************/
}

// === Move function that uses move(), not geometricBounds overwrite ===
function moveObjectAbove(mover, base) {
  var moverBounds = mover.geometricBounds;
  var baseBounds = base.geometricBounds;

  var moverHeight = moverBounds[2] - moverBounds[0];

  // Calculate new position: just above the base object
  var newTopY = baseBounds[0] - moverHeight;
  var currentTopY = moverBounds[0];

  var deltaY = newTopY - currentTopY;
  var deltaX = 0; // Keep original X position

  mover.move(undefined, [deltaX, deltaY]);
}


function getBoundsOfMyObject(myObjectToAling) {
  var myBoundsFrame = myObjectToAling.geometricBounds; // HW dreptunghi alb
  var myYF = myBoundsFrame[0];
  var myXF = myBoundsFrame[1];
  var myHF = myBoundsFrame[2];
  var myWF = myBoundsFrame[3];
  var W_obj = myWF - myXF;
  var H_obj = myHF - myYF;
  var newDim = [];
  newDim.push(myYF, myXF, myHF, myWF, W_obj, H_obj);
  return newDim;
}



// pune in bleed
function fixbleed(myObjectToAling) {
  // var doc = app.activeDocument;
  var doc = app.documents[0];
  var myBleeed_ = doc.documentPreferences.properties.documentBleedTopOffset;
  // Minimum bleed — the amount of bleed expected. To the right is a drop down list of measurement units used for all values.
  // When less than — the distance from any edge that if less than, an element should bleed.

  var edgeMM = 20;
  var bleed = myBleeed_ * 2.8346438836889; // ??? myBleeed_ e in mm??? //"5 mm"//valueAsPoints(inpMinBleed, listUnits);
  var fromEdge = edgeMM * 2.8346438836889; // "0 mm"//valueAsPoints(inpFromEdge, listUnits);


  var item = myObjectToAling;

  spread = doc.spreads[0];
  boundsSpread = [0, 0, 0, 0];
  // for (var ii = 0; ii < spread.pages.length; ii += 1) {
  boundsSpread[2] = Math.max(boundsSpread[2], spread.pages[0].bounds[2]);
  boundsSpread[3] = Math.max(boundsSpread[3], spread.pages[0].bounds[3]);
  // }
  boundsBleed = [
    -bleed,
    -bleed,
    boundsSpread[2] + bleed,
    boundsSpread[3] + bleed,
  ];
  boundsFromEdge = [
    fromEdge,
    fromEdge,
    boundsSpread[2] - fromEdge,
    boundsSpread[3] - fromEdge,
  ];

  processItem(item);

  function processItem(item) {
    boundsItem = item.geometricBounds;
    graphic = item.allGraphics[0];
    ba = [4];
    ba[0] = boundsItem[0] < boundsFromEdge[0]; // true or false
    ba[1] = boundsItem[1] < boundsFromEdge[1];
    ba[2] = boundsItem[2] > boundsFromEdge[2];
    ba[3] = boundsItem[3] > boundsFromEdge[3];
    if (!(ba[0] || ba[1] || ba[2] || ba[3])) {
      return;
    }
    boundsVisible = [
      boundsItem[0],
      boundsItem[1],
      boundsItem[2],
      boundsItem[3],
    ];
    if (graphic) {
      boundsGraphic = graphic.geometricBounds;
      boundsVisible[0] = Math.max(boundsItem[0], boundsGraphic[0]);
      boundsVisible[1] = Math.max(boundsItem[1], boundsGraphic[1]);
      boundsVisible[2] = Math.min(boundsItem[2], boundsGraphic[2]);
      boundsVisible[3] = Math.min(boundsItem[3], boundsGraphic[3]);
    }
    ba[0] = ba[0] && boundsVisible[0] > boundsBleed[0] + 0.005;
    ba[1] = ba[1] && boundsVisible[1] > boundsBleed[1] + 0.005;
    ba[2] = ba[2] && boundsVisible[2] < boundsBleed[2] - 0.005;
    ba[3] = ba[3] && boundsVisible[3] < boundsBleed[3] - 0.005;
    if (!(ba[0] || ba[1] || ba[2] || ba[3])) {
      return;
    }
    // if (item.getElements()[0] instanceof Group) {
    // 	group = item.getElements()[0];
    // 	for (var i = 0; i < group.pageItems.length; i += 1) {
    // 		if (group.pageItems[i].getElements()[0] instanceof Polygon) {
    // 			showItem(group);
    // 			alert(" Element is part of a group\nThe script cannot determine if the group bleeds correctly.Examine the group to confirm if bleed is a problem.If needed, correct the bleed manually.", title, false);
    // 			break;
    // 		}
    // 		processItem(group.pageItems[i]);
    // 	}
    // 	return;
    // }
    isColored =
      item.fillColor.name != " None " ||
      (item.strokeColor.name != " None " && item.strokeWeight > 0);
    if (!graphic && !isColored) {
      return;
    }
    // confirmChange(" Item does not bleed enough ", item, function () {
    changeItem(item);
    function changeItem(item) {
      try {
        if (graphic) {
          if (!ba[0] && boundsVisible[0] > boundsItem[0]) {
            // top
            boundsItem[0] = boundsVisible[0];
          }
          if (!ba[1] && boundsVisible[1] > boundsItem[1]) {
            // left
            boundsItem[1] = boundsVisible[1];
          }
          if (!ba[2] && boundsVisible[2] < boundsItem[2]) {
            //bottom
            boundsItem[2] = boundsVisible[2];
          }
          if (!ba[3] && boundsVisible[3] < boundsItem[3]) {
            //right
            boundsItem[3] = boundsVisible[3];
          }
          if (ba[0] && boundsGraphic[0] > boundsBleed[0]) {
            diff = boundsGraphic[0] - boundsBleed[0];
            size = boundsGraphic[2] - boundsGraphic[0];
            scale = (size + diff) / size;

            // top

            // alert("1") // sus top
            graphic.resize(
              BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
              AnchorPoint.BOTTOM_CENTER_ANCHOR,
              ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
              [scale, scale]
            );
            processItem(item);
          }
          if (ba[1] && boundsGraphic[1] > boundsBleed[1]) {
            diff = boundsGraphic[1] - boundsBleed[1];
            size = boundsGraphic[3] - boundsGraphic[1];
            scale = (size + diff) / size;

            //left

            // alert("2") // stanga
            graphic.resize(
              BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
              AnchorPoint.RIGHT_CENTER_ANCHOR,
              ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
              [scale, scale]
            );
            processItem(item);
          }
          if (ba[2] && boundsGraphic[2] < boundsBleed[2]) {
            diff = boundsBleed[2] - boundsGraphic[2];
            size = boundsGraphic[2] - boundsGraphic[0];
            scale = (size + diff) / size;

            //bottom

            // alert("3") // 3  este jos nu trebuie
            graphic.resize(
              BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
              AnchorPoint.TOP_CENTER_ANCHOR,
              ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
              [scale, scale]
            );
            processItem(item);
          }
          if (ba[3] && boundsGraphic[3] < boundsBleed[3]) {
            diff = boundsBleed[3] - boundsGraphic[3];
            size = boundsGraphic[3] - boundsGraphic[1];
            scale = (size + diff) / size;

            //right

            // alert("4") // dreapta
            graphic.resize(
              BoundingBoxLimits.GEOMETRIC_PATH_BOUNDS,
              AnchorPoint.LEFT_CENTER_ANCHOR,
              ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY,
              [scale, scale]
            );
            processItem(item);
          }
        }
        if (ba[0]) {
          //top
          boundsItem[0] = boundsBleed[0];
        }
        if (ba[1]) {
          // left
          boundsItem[1] = boundsBleed[1];
        }
        if (ba[2]) {
          // bottom
          boundsItem[2] = boundsBleed[2];
        }
        if (ba[3]) {
          //right
          boundsItem[3] = boundsBleed[3];
        }
        // item.geometricBounds = boundsItem;
      } catch (e) {
        eic = e;
        throw e;
      }
    } // });
  }
  // for (var i = 0; i < arrEroriBleed.length; i++) {
  // 	alert(arrEroriBleed[i])

  // }
  // function showItem(item) {
  // 	boundsItem = item.geometricBounds;
  // 	hItem = Number(boundsItem[2] - boundsItem[0]);
  // 	wItem = Number(boundsItem[3] - boundsItem[1]);
  // 	boundsPage = app.activeWindow.activePage.bounds;
  // 	hPage = Number(boundsPage[2] - boundsPage[0]);
  // 	wPage = Number(boundsPage[3] - boundsPage[1]);
  // 	factor = Math.min(hPage / hItem, wPage / wItem);
  // 	app.activeWindow.zoom(ZoomOptions.fitPage);
  // 	item.select();
  // 	app.activeWindow.zoomPercentage *= factor;
  // }
}
function lockLayers() {
  //********************************** lock layers + active layer */

  try {
    app.documents.item(0).layers.itemByName("id").locked ^= 1;
    app.documents.item(0).layers.itemByName("vizibil").locked ^= 1;
    app.documents.item(0).layers.itemByName("HW").locked ^= 1;
    app.documents.item(0).layers.itemByName("Panel").locked ^= 1;
    app.documents.item(0).layers.itemByName("Image").locked ^= 1;
    app.documents.item(0).layers.itemByName("DTI").locked ^= 1;
    app.documents.item(0).layers.itemByName("background").locked ^= 1;

    app.documents.item(0).layers.itemByName("Image").locked = true;
    app.documents.item(0).layers.itemByName("Image").locked = true;
    app.documents.item(0).layers.itemByName("DTI").locked = true;

    app.documents.item(0).layers.itemByName("id").visible = true;
    app.documents.item(0).layers.itemByName("vizibil").visible = true;
  } catch (e) {
    // alert(e);
  }

  //********************************** lock layers */
}

function ordoneazaPagini_dupa_Ratie() {
  //********************************** ORDONEAZA PAGINI DUPA RATIE START

  app.documents[0].documentPreferences.facingPages = false;
  var spread = app.documents[0].spreads.everyItem();
  spread.allowPageShuffle = true;
  app.documents[0].documentPreferences.allowPageShuffle = true;

  var myRatiaUnu = new Array();

  citestepaginile(myRatiaUnu);

  function citestepaginile(myRatiaUnu) {
    for (var i = 0; i < app.documents[0].pages.length; i++) {
      myPage = app.documents[0].pages.item(i);

      var b = myPage.bounds;
      var W_ = b[3] - b[1]; //
      var H_ = b[2] - b[0]; //
      var ratia = W_ / H_;
      myRatiaUnu.push(ratia);
    }
    // alert("ratia din for din functia citeste " + myRatia)
    // alert("ratia length " + myRatia.length)
    comparaRatia_simuta(myRatiaUnu);
  }

  function comparaRatia_simuta(myRatiaUnu) {
    for (var myCounter = 0; myCounter < myRatiaUnu.length - 1; myCounter++) {
      if (myRatiaUnu[myCounter] > myRatiaUnu[myCounter + 1]) {
        app.documents[0].spreads
          .item(myCounter)
          .move(
            LocationOptions.AFTER,
            app.documents[0].spreads.item(myCounter + 1)
          );
        myRatiaUnu = [];
        citestepaginile(myRatiaUnu);
      }
    }
  }

  app.documents[0].save();

  //**********************************  ORDONEAZA PAGINI DUPA RATIE END
}



function GetDataFromExcelPC_INDD(excelFilePath, sheetNumber) {
  var w = 5400;
  var h = 5400;
  // var units = "MM"
  var myDoc = createNewDoc(w, h);

  var data = getData(excelFilePath, sheetNumber);

  function getData(excelFilePath) {
    // var myExcelPath = File("g:/lucru/marcu/censhare/teste pt script/testImport.xlsx")

    var myExcelPath = File(excelFilePath);
    // var myExcelPath = excelFilePath;
    var myFrame = myDoc.textFrames.add({ geometricBounds: [0, 0, 5200, 5200] }); //app.activeDocument

    app.excelImportPreferences.properties = {
      alignmentStyle: AlignmentStyleOptions.CENTER_ALIGN,
      preserveGraphics: false,
      // rangeName
      sheetIndex: 0, //sheetNumber
      showHiddenCells: false,
      tableFormatting: TableFormattingOptions.EXCEL_UNFORMATTED_TABBED_TEXT, //TableFormattingOptions.EXCEL_FORMATTED_TABLE,
      useTypographersQuotes: false,
      // viewName: "" // asta da eroare
    };

    var myProperties = app.excelImportPreferences.properties;

    myFrame.place(myExcelPath, false, myProperties);
    // var myString = myFrame.contents + "\r" + myFrame.parentStory.insertionPoints.itemByRange(myFrame.insertionPoints[-1].index, myFrame.parentStory.insertionPoints[-1].index).contents;
    var myString = myFrame.contents;
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
    app.findGrepPreferences.findWhat = "\n";
    app.changeGrepPreferences.changeTo = " ";
    myFrame.changeGrep();
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

    var lines = myString.split("\r");
    // alert(lines.length);
    // alert(lines[0]);

    // Initialize an empty multi-dimensional array
    var multiArray = [];

    // Iterate through each line
    for (var i = 1; i < lines.length; i++) {
      // Split the line by tabs to get the elements
      var elements = lines[i].split("\t");

      // Push the elements array into the multi-dimensional array
      multiArray.push(elements);
    }

    try {
      myDoc.close(SaveOptions.no);
    } catch (e) { }
    return multiArray;
  }
  // alert(data[1][0])
  return data;
}
function createNewDoc(w, h) {
  // Create new doc.
  var d = app.documents.add();
  d.documentPreferences.facingPages = false;
  d.documentPreferences.allowPageShuffle = true;
  d.documentPreferences.documentBleedTopOffset = 0;
  d.documentPreferences.documentBleedUniformSize = true;
  // Margins to zero, single column.
  d.marginPreferences.left = 0;
  d.marginPreferences.top = 0;
  d.marginPreferences.right = 0;
  d.marginPreferences.bottom = 0;
  d.marginPreferences.columnCount = 1;
  d.marginPreferences.columnGutter = 0;
  // Set the measurement units and ruler origin.
  d.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  d.zeroPoint = [0, 0];
  // Remove excess master page.
  if (d.masterSpreads[0].pages.length > 1) {
    d.masterSpreads[0].pages[1].remove();
  }
  // Set master page margins.
  d.masterSpreads[0].pages[0].marginPreferences.left = 0;
  d.masterSpreads[0].pages[0].marginPreferences.top = 0;
  d.masterSpreads[0].pages[0].marginPreferences.right = 0;
  d.masterSpreads[0].pages[0].marginPreferences.bottom = 0;
  d.masterSpreads[0].pages[0].marginPreferences.columnCount = 1;
  d.masterSpreads[0].pages[0].marginPreferences.columnGutter = 0;
  // Set page dimensions.
  d.documentPreferences.pageWidth = w;
  d.documentPreferences.pageHeight = h;
  // Return new doc.
  return d;
}

// Function to remove duplicates
function removeDuplicates(array) {
  var uniqueArray = [];

  for (var i = 0; i < array.length; i++) {
    var currentValue = array[i];
    var isUnique = true;

    for (var j = 0; j < uniqueArray.length; j++) {
      if (currentValue === uniqueArray[j]) {
        isUnique = false;
        break;
      }
    }

    if (isUnique) {
      uniqueArray.push(currentValue);
    }
  }
  return uniqueArray;
}

//**********************************  end succes remove duplicates

function find_files(dir, mask_array) {
  var arr = [];
  for (var i = 0; i < mask_array.length; i++) {
    arr = arr.concat(Folder(dir).getFiles("*" + mask_array[i]));
  }
  return arr;
}

function padWithZeros(number, length) {
  var str = number.toString();
  var nrLength = length.toString();
  while (str.length < nrLength.length) {
    str = "0" + str;
  }
  return str;
}

function writeReportsErrors(arrErrors, pathofmaster) {
  for (var a = 0; a < arrErrors.length; a++) {
    var message = arrErrors[a];
    // alert("message: " + message)
    var filename = "_Errors_report" + dateAnZiOraMin + "_" + ".txt"; // merge!

    var file = new File(pathofmaster + "/" + filename);

    // alert("file: " + filename + " filename: " + filename + " calea e ")

    file.encoding = "UTF-8";

    if (file.exists) {
      file.open("e");
      file.seek(0, 2);
    } else {
      file.open("w");
    }

    // file.open('w');
    file.write(message + "\r");
    file.close();
  }
  alert("Some errors found, check Error report!");
}
// Function to find .xlsx files in a folder
function findXLSXFiles(folderPath) {
  var folder = new Folder(folderPath); // Create a Folder object
  if (!folder.exists) {
    alert("The folder does not exist.");
    return [];
  }

  // Get all the files in the folder that end with .xlsx
  var files = folder.getFiles(function (file) {
    return file instanceof File && file.name.match(/\.xlsx$/i);
  });

  return files; // Return an array of .xlsx files
}



///////---- hw dti si linia





function aliniere_zece_la_suta(myParent, mychild) { // HW ul de ex, sau dti cand e singur


  var myPage = app.documents[0].pages.item(0);
  var b_pgebounds = myPage.bounds;

  var w_t = b_pgebounds[3] - b_pgebounds[1];
  var h_t = b_pgebounds[2] - b_pgebounds[0];

  var m_left = myPage.marginPreferences.left; //mmyX2
  var m_right = myPage.marginPreferences.right; //mmyX1
  var m_top = myPage.marginPreferences.top; //mmyy2
  var m_bottom = myPage.marginPreferences.bottom; //mmyy1
  var w = w_t - 2 * m_left; // w vizibil
  var h = h_t - 2 * m_top; // h vizibil

  //**********************************    HW     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  var HW_h_final = h * HwSize; //0.1); // 10% hw
  // var SAIZECI_Hwai = HW_h_final * 0.6; // 60% din inaltime lui h
  var Y_HW_final = m_top + h - HW_h_final;
  var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;


  var myHW = myParent; // HW
  var myHWai = mychild;


  // var myHW = app.documents[0].rectangles.itemByName("hwul"); // HW
  // var myHWai = app.documents[0].rectangles.itemByName("hwul").rectangles.itemByName("hwaiul"); // ai ul din HW!!!

  myHW.redefineScaling([1, 1]);
  myHWai.redefineScaling([1, 1]);

  myHW.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left]; //succes
  app.documents[0].align(
    myHWai,
    AlignOptions.HORIZONTAL_CENTERS,
    AlignDistributeBounds.MARGIN_BOUNDS
  );

  myHW.fit(FitOptions.CENTER_CONTENT);


  // var myBoundsw = myHW.geometricBounds; // HW dreptunghi alb
  // var myYw = myBoundsw[0];
  // var myXw = myBoundsw[1];
  // var myHw = myBoundsw[2];
  // var myWw = myBoundsw[3];
  var W_hw = myHW.geometricBounds[3] - myHW.geometricBounds[1];
  // var H_hw = myHW.geometricBounds[2] - myHW.geometricBounds[0];

  // var myBoundsw = myHWai.geometricBounds; // HW dreptunghi alb
  // var myYwai = myBoundsw[0];
  // var myXwai = myBoundsw[1];
  // var myHwai = myBoundsw[2];
  // var myWwai = myBoundsw[3];
  var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
  var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];
  var obj = myHWai;
  var anchor = AnchorPoint.CENTER_ANCHOR;
  // var ow = obj.geometricBounds[3] - obj.geometricBounds[1];
  // var oh = obj.geometricBounds[2] - obj.geometricBounds[0];

  if (w > h) {
    // landscape

    var pw = (HW_h_final * 0.7) / H_hwai;
    var ph = pw;
    var matrix = app.transformationMatrices.add({
      horizontalScaleFactor: pw,
      verticalScaleFactor: ph,
    });
    obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }
  if (w < h) {
    myHW.redefineScaling([1, 1]);
    myHWai.redefineScaling([1, 1]);
    var pw = (W_hw * 0.8) / W_hwai;
    var ph = pw;
    var matrix = app.transformationMatrices.add({
      horizontalScaleFactor: pw,
      verticalScaleFactor: ph,
    });
    obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }

  // var myBoundsw = myHWai.geometricBounds; // HW dreptunghi alb
  // var myYwai = myBoundsw[0];
  // var myXwai = myBoundsw[1];
  // var myHwai = myBoundsw[2];
  // var myWwai = myBoundsw[3];
  var W_hwai = myHWai.geometricBounds[3] - myHWai.geometricBounds[1];
  var H_hwai = myHWai.geometricBounds[2] - myHWai.geometricBounds[0];

  if (W_hwai > w - w * 0.1) {
    // portrait //W_hwai>w-(w*0.1)

    myHW.redefineScaling([1, 1]);
    myHWai.redefineScaling([1, 1]);
    var pw = (W_hw * 0.8) / W_hwai;
    var ph = pw;
    var matrix = app.transformationMatrices.add({
      horizontalScaleFactor: pw,
      verticalScaleFactor: ph,
    });
    obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }

  myHWai.fit(FitOptions.CENTER_CONTENT);
  ////!! modificare la fit!!
  myHWai.fit(FitOptions.frameToContent);

  myHW.fit(FitOptions.CENTER_CONTENT);

  myHW.geometricBounds = [
    Y_HW_final_latotal,
    -14.17322835,
    h_t + 14.17322835,
    w_t + 14.17322835,
  ];

  //**********************************     HW    end  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



}


function aliniere_cinci_la_suta(myParent, mychild) { // fostul dti dti40 cand e si HW prezent

  var myPage = app.documents[0].pages.item(0);
  var b_pgebounds = myPage.bounds;

  var w_t = b_pgebounds[3] - b_pgebounds[1];
  var h_t = b_pgebounds[2] - b_pgebounds[0];

  var m_left = myPage.marginPreferences.left; //mmyX2
  var m_right = myPage.marginPreferences.right; //mmyX1
  var m_top = myPage.marginPreferences.top; //mmyy2
  var m_bottom = myPage.marginPreferences.bottom; //mmyy1
  var w = w_t - 2 * m_left; // w vizibil
  var h = h_t - 2 * m_top; // h vizibil

  //**********************************    HW     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  var HW_h_final = h * HwSize; //0.1); // 10% hw
  // var SAIZECI_Hwai = HW_h_final * 0.6; // 60% din inaltime lui h
  var Y_HW_final = m_top + h - HW_h_final;
  var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;

  var Hdt_final = (h - HW_h_final) * DtiSize; //0.04); // 4% dt din h minus hw
  var Y_DT_final = m_top + h - (HW_h_final + Hdt_final);
  // var Y_DT_final_latotal = h - Hdt_final + (h_t - h) / 2;


  var myDT = myParent
  var myDTai = mychild



  // var myDT = app.documents[0].rectangles.itemByName("dt4"); // HW
  // var myDTai = app.documents[0].rectangles.itemByName("dt4").rectangles.itemByName("dt4ai"); //
  // var mylinedthw = app.documents[0].graphicLines.itemByName("linedthw"); //

  myDT.redefineScaling([1, 1]);
  myDTai.redefineScaling([1, 1]);

  myDT.geometricBounds = [
    Y_DT_final,
    -14.174,
    Y_DT_final + Hdt_final,
    w_t + 14.174,
  ]; //

  myDT.fit(FitOptions.CENTER_CONTENT);
  ////!! modificare la fit!!
  myDTai.fit(FitOptions.frameToContent);


}


function aliniere_linia(myLine) {

  var myPage = app.documents[0].pages.item(0);
  var b_pgebounds = myPage.bounds;

  var w_t = b_pgebounds[3] - b_pgebounds[1];
  var h_t = b_pgebounds[2] - b_pgebounds[0];

  var m_left = myPage.marginPreferences.left; //mmyX2
  var m_right = myPage.marginPreferences.right; //mmyX1
  var m_top = myPage.marginPreferences.top; //mmyy2
  var m_bottom = myPage.marginPreferences.bottom; //mmyy1
  var w = w_t - 2 * m_left; // w vizibil
  var h = h_t - 2 * m_top; // h vizibil

  //**********************************    HW     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  var HW_h_final = h * HwSize; //0.1); // 10% hw
  // var SAIZECI_Hwai = HW_h_final * 0.6; // 60% din inaltime lui h
  var Y_HW_final = m_top + h - HW_h_final;
  var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;

  var Hdt_final = (h - HW_h_final) * DtiSize; //0.04); // 4% dt din h minus hw
  var Y_DT_final = m_top + h - (HW_h_final + Hdt_final);
  // var Y_DT_final_latotal = h - Hdt_final + (h_t - h) / 2;

  var mylinedthw = myLine
  mylinedthw.redefineScaling([1, 1]);

  mylinedthw.geometricBounds = [
    Y_DT_final + Hdt_final,
    -14.174,
    Y_DT_final + Hdt_final,
    w_t + 14.174,
  ];
  mylinedthw.strokeWeight = 1;
}