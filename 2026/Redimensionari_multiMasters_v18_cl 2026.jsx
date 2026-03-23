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

last update: 18 martie 2026

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


// aliniere expand to bleed
//    1
//  _____
// 4|   |2
//  _____
//   3

//-------

if (!app.documents.length > 0) {
  alert("Please open any master file!");
  exit();
}

app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.generalPreferences.pageNumbering = PageNumberingOptions.absolute;
app.scriptPreferences.enableRedraw = true;
app.layoutWindows[0].transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var HwSize = 10 / 100;
var DtiSize = 5 / 100;
var nameOfFoldeResult = "_result";
var substractFromBottom;
var arrErrors = [];

var d = new Date();
var dateAnZiOraMin = "_" + d.getDate() + "_" + (d.getMonth() + 1) + "_" + d.getFullYear() + "_" + d.getHours() + "_" + d.getMinutes() + "_" + d.getSeconds();

var myFile_calea = app.documents[0].filePath;
var myFileName = app.documents[0].name + "";
var myFileName0 = myFileName.substr(0, myFileName.lastIndexOf("."));
app.documents[0].close(SaveOptions.no);

var pathFileMaster = myFile_calea;
var osconnector = File.fs == "Windows" ? "\\" : "/";

var excelFilePathArr = find_files(pathFileMaster, [".xlsx"]);
if (excelFilePathArr.length === 1) {
  excelFilePath = excelFilePathArr[0];
} else {
  alert("Please keep one excel file in same folder like master files!\nOr close the opened excel file!");
  exit();
}

var dataList = GetDataFromExcelPC_INDD(excelFilePath);

dataList = checkMyData(dataList);
function checkMyData(arr) {
  function isValid(val) {
    if (val === undefined) return false;
    var v = val.replace(",", ".");
    return !isNaN(parseFloat(v)) && isFinite(v);
  }
  var cleaned = [];
  for (var i = 0; i < arr.length; i++) {
    var row = arr[i];
    var firstVal = row[0];
    var allSame = true;
    for (var k = 1; k < row.length; k++) {
      if (row[k] !== firstVal) { allSame = false; break; }
    }
    if (!allSame) {
      cleaned.push(row);
      for (var j = 2; j < 6; j++) {
        if (!isValid(row[j])) {
          arrErrors.push("A dimension is " + (row[j] === undefined ? "not defined" : "not a number") + " at line " + (i + 2) + " column " + (j + 1));
        }
      }
    }
  }
  return cleaned;
}

if (arrErrors.length > 0) {
  writeReportsErrors(arrErrors, pathFileMaster);
  exit();
}

var numberOfFiles = dataList.length;
var arrMasterName1 = [];
var inddPattern = /.indd/;
for (var i = 0; i < dataList.length; i++) {
  if (inddPattern.test(dataList[i][7])) arrMasterName1.push(dataList[i][7]);
}

var arrMasterName = removeDuplicates(arrMasterName1);



for (var i = 0; i < arrMasterName.length; i++) {
  var filePath = pathFileMaster + "/" + arrMasterName[i];
  if (!new File(filePath).exists) {
    alert("The master file does not exist: " + filePath);
    exit();
  }
  app.open(File(filePath), false);
  ordoneazaPagini_dupa_Ratie();
  app.documents[0].save();
  app.documents[0].close(SaveOptions.no);
}


for (var i = 0; i < arrMasterName.length; i++) {
  var pfPath = pathFileMaster + "/" + arrMasterName[i];
  if (!new File(pfPath).exists) continue;
  app.open(File(pfPath), false);
  var pfDoc = app.documents[0];
  for (var pp = 0; pp < pfDoc.pages.length; pp++) {
    var pfItems = pfDoc.pages.item(pp).allPageItems;
    for (var pi = 0; pi < pfItems.length; pi++) {
      var pfLabel = pfItems[pi].label;
      if ((pfLabel === "Health_Warning" || pfLabel === "DeTeIi") && pfItems[pi].rectangles.length === 0) {
        arrErrors.push("Label \"" + pfLabel + "\" has no child rectangle (linked file must be in a rectangle inside a rectangle) - page " + (pp + 1) + " in " + arrMasterName[i]);
      }
    }
  }
  pfDoc.close(SaveOptions.no);
}
if (arrErrors.length > 0) {
  writeReportsErrors(arrErrors, pathFileMaster);
  exit();
}




function CreateProgressBar() {
  var w = new Window("window", "Redimensionare");
  w.pb = w.add("progressbar", [12, 12, 1250, 25], 0, undefined);
  w.st = w.add("statictext");
  w.st.bounds = [0, 0, 1200, 20];
  w.st.alignment = "left";
  return w;
}

var progressWin = CreateProgressBar();
progressWin.show();
progressWin.update();
progressWin.pb.minvalue = 0;
progressWin.pb.maxvalue = Number(numberOfFiles);
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myDocument, vizibil_W_mm, vizibil_H_mm, total_W_mm, total_H_mm, masterNameFile, varInfoCol4;

for (var r = 0; r < dataList.length; r++) {
  var KVname = dataList[r][7].replace(".indd", "");
  var NameOfMyCampain = dataList[r][8];
  var material_places = cleanMydata(dataList[r][1]);
  if (material_places.length <= 2) material_places = "";

  var CountNumber = padWithZeros(r + 1, numberOfFiles);
  var unitsofdimensions = dataList[r][6];

  var dimensionsname;
  if (Number(dataList[r][4]) == Number(dataList[r][2]) && Number(dataList[r][5]) == Number(dataList[r][3])) {
    dimensionsname = Number(dataList[r][4]) + "x" + Number(dataList[r][5]);
  } else {
    dimensionsname = Number(dataList[r][4]) + "x" + Number(dataList[r][5]) + "_" + Number(dataList[r][2]) + "x" + Number(dataList[r][3]);
  }

  var finalFileName = CountNumber + "_" + NameOfMyCampain + "_" + material_places + "__" + dimensionsname + "_" + unitsofdimensions + "_KV_" + KVname;

  varInfoCol4 = dataList[r][6];
  masterNameFile = dataList[r][7];
  var idNumar = dataList[r][0];

  progressWin.pb.value = r;
  progressWin.st.text = "Processing file: \r" + finalFileName + "\rMaster: " + masterNameFile + "\r(" + r + " / " + numberOfFiles + ")";
  progressWin.update();
  // alert("Processing file: \r" + finalFileName + "\rMaster: " + masterNameFile + "\r(" + r + " / " + numberOfFiles + ")");
  app.open(File(pathFileMaster + osconnector + masterNameFile), false);

  var myRatia = [];
  for (var myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
    var b = app.documents[0].pages.item(myCounter).bounds;
    myRatia.push((b[3] - b[1]) / (b[2] - b[0]));
  }

  myDocument = app.documents[0];
  var myPage_length = myDocument.pages.length;
  var myPages = myDocument.pages;

  myDocument.layers.everyItem().locked = false;

  try {
    myDocument.layers.add({ name: "vizibil", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    myDocument.layers.add({ name: "id", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    myDocument.layers.itemByName("vizibil").move(LocationOptions.AT_BEGINNING);
    myDocument.layers.itemByName("id").move(LocationOptions.AT_BEGINNING);
  } catch (e) { }

  var units = applyUnits(myDocument, dataList[r][6]);
  var changeUnitsOfDocument = dataList[r][6] || "mm";

  var primu_W_mm = String(dataList[r][2]).replace(/,/g, ".");
  var primuH_mm = String(dataList[r][3]).replace(/,/g, ".");
  var aldoilea_W_mm = String(dataList[r][4]).replace(/,/g, ".");
  var aldoilea_H_mm = String(dataList[r][5]).replace(/,/g, ".");

  var vizibil_W, vizibil_H, total_W, total_H;
  if (primu_W_mm * primuH_mm < aldoilea_W_mm * aldoilea_H_mm) {
    vizibil_W = Number(primu_W_mm * units);
    vizibil_H = Number(primuH_mm * units);
    total_W = Number(aldoilea_W_mm * units);
    total_H = Number(aldoilea_H_mm * units);
    vizibil_W_mm = dataList[r][2];
    vizibil_H_mm = dataList[r][3];
    total_W_mm = dataList[r][4];
    total_H_mm = dataList[r][5];
  } else {
    vizibil_W = Number(aldoilea_W_mm * units);
    vizibil_H = Number(aldoilea_H_mm * units);
    total_W = Number(primu_W_mm * units);
    total_H = Number(primuH_mm * units);
    vizibil_W_mm = dataList[r][4];
    vizibil_H_mm = dataList[r][5];
    total_W_mm = dataList[r][2];
    total_H_mm = dataList[r][3];
  }

  var pagDeExtras = findClosestElement(myRatia, vizibil_W / vizibil_H);

  for (var i = myPage_length - 1; i >= 0; i--) {
    if (i !== pagDeExtras) myPages[i].remove();
  }

  var numeRatieFolder = Number(myRatia[pagDeExtras]).toFixed(3).replace(/\./g, "_");

  try {
    new Folder(pathFileMaster + "/" + nameOfFoldeResult).create();
  } catch (e) { }

  scaleDocument(vizibil_W, vizibil_H);
  resizeDocument(total_W, total_H);
  idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar);
  addMyGuides();
  aliniereElemente(vizibil_W, vizibil_H, total_W, total_H);
  lockLayers();

  applyUnits(myDocument, changeUnitsOfDocument);

  var savedoc = myDocument.save(File(pathFileMaster + "/" + nameOfFoldeResult + "/" + finalFileName + "_.indd"));
  savedoc.close(SaveOptions.no);
}

progressWin.close();
alert("Job done!\nFiles are in " + nameOfFoldeResult + " folder");

function applyUnits(doc, unitStr) {
  var unitMap = {
    "mm": { h: MeasurementUnits.MILLIMETERS, v: MeasurementUnits.MILLIMETERS, factor: 2.83464567 },
    "cm": { h: MeasurementUnits.CENTIMETERS, v: MeasurementUnits.CENTIMETERS, factor: 28.3464567 },
    "px": { h: MeasurementUnits.PIXELS, v: MeasurementUnits.PIXELS, factor: 1 },
    "inch": { h: MeasurementUnits.INCHES, v: MeasurementUnits.INCHES, factor: 72 }
  };
  var u = unitMap[unitStr] || unitMap["mm"];
  doc.viewPreferences.horizontalMeasurementUnits = u.h;
  doc.viewPreferences.verticalMeasurementUnits = u.v;
  return u.factor;
}

function findClosestElement(arr, target) {
  var closest = 0;
  var minDiff = Math.abs(target - arr[0]);
  for (var p = 1; p < arr.length; p++) {
    var diff = Math.abs(target - arr[p]);
    if (diff < minDiff) { minDiff = diff; closest = p; }
  }
  return closest;
}

function cleanMydata(raw) {
  // return raw.replace(/[\/\.\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");
  return raw.replace(/[\/\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");
}

function scaleDocument(vizibil_W, vizibil_H) {
  var doc = app.documents[0];
  doc.zeroPoint = [0, 0];
  doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  var page = doc.pages[0];
  page.marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
  page.layoutRule = LayoutRuleOptions.SCALE;
  page.resize(CoordinateSpaces.SPREAD_COORDINATES, AnchorPoint.CENTER_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [vizibil_W, vizibil_H]);
}

function resizeDocument(total_W, total_H) {
  var page = app.documents[0].pages[0];
  page.marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
  page.layoutRule = LayoutRuleOptions.OFF;
  page.resize(CoordinateSpaces.SPREAD_COORDINATES, AnchorPoint.CENTER_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [total_W, total_H]);
}

function idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar) {
  try {
    app.documents[0].activeLayer = app.documents[0].layers.itemByName("id");
    app.documents.item(0).layers.itemByName("id").locked = false;
    app.documents.item(0).layers.itemByName("vizibil").locked = false;
  } catch (e) { }

  var myPage = app.documents[0].pages.item(0);
  var textFrame_info = myPage.textFrames.add();
  textFrame_info.geometricBounds = [0, total_W + 50, 300, total_W + 500];
  textFrame_info.contents = "W vizibil = " + vizibil_W_mm + "\nH vizibil = " + vizibil_H_mm + "\nW total = " + total_W_mm + "\nH total = " + total_H_mm + "\n - " + varInfoCol4 + "\n - " + masterNameFile;
  textFrame_info.label = "Info";
  textFrame_info.parentStory.paragraphs.item(0).paragraphs.everyItem().properties = {
    appliedFont: app.fonts.item("Times"),
    fontStyle: "Regular",
    pointSize: 25,
    leading: 30,
    fillColor: "Black",
  };

  try {
    appliedLanguage = "English: USA";
    textFrame_info.fit(FitOptions.FRAME_TO_CONTENT);
    textFrame_info.move([total_W + 20, 0]);
  } catch (e) { }

  var m_top = (total_H - vizibil_H) / 2;
  var m_left = (total_W - vizibil_W) / 2;
  myPage.marginPreferences.properties = { top: m_top, left: m_left, right: m_left, bottom: m_top };

  app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
  var myItem_vizibil = myPage.rectangles.add({
    geometricBounds: [m_top, m_left, m_top + vizibil_H, m_left + vizibil_W],
  });

  var myColorforvizibil = app.documents[0].colors.add({
    name: "colorforvisible",
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
  myItem_vizibil.transparencySettings.blendingSettings.opacity = 100;
  myItem_vizibil.transparencySettings.blendingSettings.blendMode = BlendMode.NORMAL;
}

function addMyGuides() {
  var myPage = app.documents[0].pages.item(0);
  var mp = myPage.marginPreferences;
  var b = myPage.bounds;
  var W_ = b[3] - b[1];
  var H_ = b[2] - b[0];
  app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: W_ - mp.right });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: mp.left });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: mp.top });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: H_ - mp.bottom });
}

function aliniereElemente(vizibil_W, vizibil_H, total_W, total_H) {
  var myPage = app.documents[0].pages.item(0);
  var pageItems = myPage.allPageItems;
  var m_top = myPage.marginPreferences.properties.top;
  var m_left = myPage.marginPreferences.properties.left;
  var m_right = myPage.marginPreferences.properties.right;
  var m_bottom = myPage.marginPreferences.properties.bottom;
  var b = myPage.bounds;
  var W_ = b[3] - b[1];
  var H_ = b[2] - b[0];
  var H_vizibil = H_ - (m_top + m_bottom);
  var W_vizibil = W_ - (m_left + m_right);
  var doc = app.documents[0];
  var items = doc.allPageItems;

  var foundLabelHW = false, foundLabelDTI = false;
  for (var i = 0; i < items.length; i++) {
    var lbl = items[i].label;
    if (!lbl) continue;
    if (lbl === "Health_Warning") foundLabelHW = true;
    if (lbl === "DeTeIi") foundLabelDTI = true;
    if (foundLabelHW && foundLabelDTI) break;
  }

  if (foundLabelHW && foundLabelDTI) {
    substractFromBottom = H_vizibil * HwSize + (H_vizibil - H_vizibil * HwSize) * DtiSize;
  } else if (foundLabelHW) {
    substractFromBottom = H_vizibil * HwSize;
  } else {
    substractFromBottom = 0;
  }

  var douaZeciLaSuta = H_vizibil * 0.2;
  var key_optezeci_portret = W_vizibil * 0.8;
  var key_optezecisicinci_landscape = (H_vizibil - H_vizibil * 0.1) * 0.85;

  var myObjectList = [];
  var validTypes = { "Rectangle": 1, "Oval": 1, "Polygon": 1, "TextFrame": 1, "Group": 1, "Button": 1, "GraphicLine": 1 };
  for (var myCounter = 0; myCounter < pageItems.length; myCounter++) {
    if (validTypes[pageItems[myCounter].constructor.name]) myObjectList.push(pageItems[myCounter]);
  }

  var labelActions = {
    "Health_Warning": function (obj) {
      var child = obj.rectangles.length > 0 ? obj.rectangles.item(0) : null;
      if (!child) { alert("Health_Warning label found but child rectangle is missing in: " + masterNameFile); return; }
      aliniere_zece_la_suta(obj, child);
    },
    "DeTeIi": function (obj) {
      var child = obj.rectangles.length > 0 ? obj.rectangles.item(0) : null;
      if (!child) { alert("DeTeIi label found but child rectangle is missing in: " + masterNameFile); return; }
      aliniere_cinci_la_suta(obj, child);
    },
    "linia": function (obj) { aliniere_linia(obj); },
    "alignTo_Visible_Bottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
    },
    "alignTo_Visible_Corner_leftBottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
      doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Corner_rightBottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
      doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "cheia": function (obj) {
      var anchor = AnchorPoint.CENTER_ANCHOR;
      var H_ch = obj.geometricBounds[2] - obj.geometricBounds[0];
      var pw, matrix;
      if (H_ > W_) {
        pw = key_optezeci_portret / H_ch;
        matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
        doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        doc.align(obj, AlignOptions.VERTICAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        var bn = getBoundsOfMyObject(obj);
        obj.move([bn[1], (H_ - m_bottom - H_vizibil * 0.1) / 2 - bn[5] / 2]);
      } else {
        pw = key_optezecisicinci_landscape / H_ch;
        matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
        doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        doc.align(obj, AlignOptions.VERTICAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        var bn = getBoundsOfMyObject(obj);
        obj.move([bn[1], H_ - m_bottom - H_vizibil * 0.2 - bn[5]]);
      }
      var bn = getBoundsOfMyObject(obj);
      myPage.guides.add(doc.layers.itemByName("Image"), { orientation: HorizontalOrVertical.horizontal, location: (bn[2] - bn[0]) / 2 + bn[0] });
    },
    "douzecilasuta": function (obj) {
      var anchor = AnchorPoint.CENTER_ANCHOR;
      var H_dz = obj.geometricBounds[2] - obj.geometricBounds[0];
      var pw = douaZeciLaSuta / H_dz;
      var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
      obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
      doc.align(obj, AlignOptions.BOTTOM_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      var bn = getBoundsOfMyObject(obj);
      myPage.guides.add(doc.layers.itemByName("Image"), { orientation: HorizontalOrVertical.horizontal, location: H_ - bn[5] - m_bottom });
    },
    "alignToPageBottom": function (obj) { doc.align(obj, AlignOptions.BOTTOM_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageLeft": function (obj) { doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageRight": function (obj) { doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageTop": function (obj) { doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignTo_Visible_Corner_leftTop": function (obj) {
      doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Corner_RightTop": function (obj) {
      doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Left": function (obj) { doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignTo_Visible_Right": function (obj) { doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignTo_Visible_Top": function (obj) { doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignHorizontalCenter": function (obj) { doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignVerticalCentersMinusHW": function (obj) {
      obj.move([myXF, (H_ - H_ * HwSize) / 2 - H_obj / 2]);
    },
    "alignedTo_key_first": function (obj) {

      var secondKeyObject = obj;

      var group = null;
      var rect = null;
      //alignTo_Visible_Bottom-HW+alignedTo_key_first
      // Find the group by name
      try {
        for (var i = 0; i < doc.allPageItems.length; i++) {
          if (doc.allPageItems[i].constructor.name === "Group" &&
            doc.allPageItems[i].name === "myGroup") {
            group = doc.allPageItems[i];
            break;
          }
        }

        if (group !== null) {
          // Find the rectangle by name inside the group
          for (var j = 0; j < group.pageItems.length; j++) {
            if (group.pageItems[j].name === "firstreferenceobject") {
              rect = group.pageItems[j];
              break;
            }
          }
        }

        if (rect !== null) {
          // app.activeDocument.select(rect);
          var firstKeyObject = rect;
          // alert("Rectangle selected: " + rect.name);
        } else {
          // alert("Rectangle not found.");
        }
        var mS = [firstKeyObject, secondKeyObject];
        doc.align(mS, AlignOptions.LEFT_EDGES, AlignDistributeBounds.KEY_OBJECT, mS[0]);
      } catch (e) {
        // alert("Error finding group: " + e.message);
      }

    },
    "expandToVizibil": function (obj) {
      obj.geometricBounds = [m_top, m_left, H_ - m_bottom, W_ - m_right];
    },
    "expandToBleed": function (obj) {
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, -bl, H_ + bl, W_ + bl];
    },
    "expandToBleed123": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, bn[1], H_ + bl, W_ + bl];
    },
    "expandToBleed234": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], -bl, H_ + bl, W_ + bl];
    },
    "expandToBleed14": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, -bl, bn[2], bn[3]];
    },
    "expandToBleed23": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], bn[1], H_ + bl, W_ + bl];
    },
    "expandToBleed42": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], -bl, bn[2], W_ + bl];
    },
    "smartBleed": function (obj) {
      if (!(obj instanceof Rectangle)) return;
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      var bp = myPage.bounds;
      var Wp = bp[3] - bp[1], Hp = bp[2] - bp[0];
      var gb = obj.geometricBounds;
      var Y1 = gb[0], X1 = gb[1], Y2 = gb[2], X2 = gb[3];
      obj.fit(FitOptions.frameToContent);
      var gf = obj.geometricBounds;
      var Y1f = gf[0], X1f = gf[1], Y2f = gf[2], X2f = gf[3];
      var caseY1 = Y1f < -bl ? -bl : (Y1f > -bl && Y1f < Hp * 0.05) ? -bl : Y1;
      var caseX1 = X1f < -bl ? -bl : (X1f > -bl && X1f < Wp * 0.05) ? -bl : X1;
      var caseY2 = Y2f > Hp + bl ? Hp + bl : (Y2f < Hp + bl && Y2f > Hp - Hp * 0.05) ? Hp + bl : Y2;
      var caseX2 = X2f > Wp + bl ? Wp + bl : (X2f < Wp + bl && X2f > Wp - Wp * 0.05) ? Wp + bl : X2;
      obj.geometricBounds = [caseY1, caseX1, caseY2, caseX2];
    },
    // if "hl_centerToLeft" is present, the geometric bounds of the object will pe modified. the margin of right of rectangle will be increased to the right margin of the page, and the left margin will stay in the same place, after that the content of the rectangle will be aligned to the center of the geometric bounds. This is useful for Headline that need to be aligned to center between left margin of rectangle and right margin of the page.
    "hl_centerToLeft": function (obj) {
      var gb = obj.geometricBounds; // [Y1, X1, Y2, X2]
      var Y1 = gb[0], X1 = gb[1], Y2 = gb[2];
      var newX2 = W_ - m_right;
      obj.geometricBounds = [Y1, X1, Y2, newX2];
      obj.fit(FitOptions.CENTER_CONTENT);
    }
  };


  // if (!(obj instanceof Rectangle)) return;
  // var bp = myPage.bounds;
  // var Wp = bp[3] - bp[1];
  // var gb = obj.geometricBounds;
  // var Y1 = gb[0], X1 = gb[1], Y2 = gb[2], X2 = gb[3];
  // var w_obj = X2 - X1;
  // var newX2 = Wp - m_right;
  // var newX1 = newX2 - w_obj;
  // obj.geometricBounds = [Y1, newX1, Y2, newX2];
  // obj.fit(FitOptions.CENTER_CONTENT);

  // Build a flat list of {obj, label} pairs, deferring "alignedTo_key_first" to the end
  var deferredList = [];

  for (var i = 0; i < myObjectList.length; i++) {
    var myObjectToAling = myObjectList[i];
    myObjectToAling.redefineScaling([1, 1]);

    var myBoundsFrame = myObjectList[i].geometricBounds;
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
        var trimmedLabel = labels[j].replace(/^\s+|\s+$/g, "");
        if (!labelActions.hasOwnProperty(trimmedLabel)) continue;
        if (trimmedLabel === "alignedTo_key_first") {
          deferredList.push({ obj: myObjectToAling, action: labelActions[trimmedLabel] });
        } else {
          labelActions[trimmedLabel](myObjectToAling);
        }
      }
    }
  }

  // Execute deferred "alignedTo_key_first" actions last
  for (var d = 0; d < deferredList.length; d++) {
    deferredList[d].action(deferredList[d].obj);
  }
}


function getBoundsOfMyObject(obj) {
  var b = obj.geometricBounds;
  return [b[0], b[1], b[2], b[3], b[3] - b[1], b[2] - b[0]];
}

function getPageMetrics() {
  var myPage = app.documents[0].pages.item(0);
  var b = myPage.bounds;
  var w_t = b[3] - b[1];
  var h_t = b[2] - b[0];
  var m_left = myPage.marginPreferences.left;
  var m_right = myPage.marginPreferences.right;
  var m_top = myPage.marginPreferences.top;
  var m_bottom = myPage.marginPreferences.bottom;
  return { w_t: w_t, h_t: h_t, m_left: m_left, m_right: m_right, m_top: m_top, m_bottom: m_bottom, w: w_t - 2 * m_left, h: h_t - 2 * m_top };
}

function lockLayers() {
  try {
    var doc = app.documents.item(0);
    var names = ["id", "vizibil", "HW", "Panel", "Image", "DTI", "background"];
    for (var i = 0; i < names.length; i++) doc.layers.itemByName(names[i]).locked ^= 1;
    doc.layers.itemByName("Image").locked = true;
    doc.layers.itemByName("DTI").locked = true;
    doc.layers.itemByName("id").visible = true;
    doc.layers.itemByName("vizibil").visible = true;
  } catch (e) { }
}

function ordoneazaPagini_dupa_Ratie() {
  var doc = app.documents[0];
  doc.documentPreferences.facingPages = false;
  doc.spreads.everyItem().allowPageShuffle = true;
  doc.documentPreferences.allowPageShuffle = true;

  function citestepaginile(arr) {
    for (var i = 0; i < doc.pages.length; i++) {
      var b = doc.pages.item(i).bounds;
      arr.push((b[3] - b[1]) / (b[2] - b[0]));
    }
    comparaRatia_simuta(arr);
  }

  function comparaRatia_simuta(arr) {
    for (var myCounter = 0; myCounter < arr.length - 1; myCounter++) {
      if (arr[myCounter] > arr[myCounter + 1]) {
        doc.spreads.item(myCounter).move(LocationOptions.AFTER, doc.spreads.item(myCounter + 1));
        citestepaginile([]);
        return;
      }
    }
  }

  citestepaginile([]);
  doc.save();
}

function GetDataFromExcelPC_INDD(excelFilePath) {
  var myDoc = createNewDoc(5400, 5400);
  var myFrame = myDoc.textFrames.add({ geometricBounds: [0, 0, 5200, 5200] });

  app.excelImportPreferences.properties = {
    alignmentStyle: AlignmentStyleOptions.CENTER_ALIGN,
    preserveGraphics: false,
    sheetIndex: 0,
    showHiddenCells: false,
    tableFormatting: TableFormattingOptions.EXCEL_UNFORMATTED_TABBED_TEXT,
    useTypographersQuotes: false,
  };

  myFrame.place(File(excelFilePath), false, app.excelImportPreferences.properties);
  var myString = myFrame.contents;

  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
  app.findGrepPreferences.findWhat = "\n";
  app.changeGrepPreferences.changeTo = " ";
  myFrame.changeGrep();
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

  var lines = myString.split("\r");
  var multiArray = [];
  for (var i = 1; i < lines.length; i++) multiArray.push(lines[i].split("\t"));
  // exit()
  try { myDoc.close(SaveOptions.no); } catch (e) { }
  return multiArray;
}

function createNewDoc(w, h) {
  var d = app.documents.add();
  d.documentPreferences.facingPages = false;
  d.documentPreferences.allowPageShuffle = true;
  d.documentPreferences.documentBleedTopOffset = 0;
  d.documentPreferences.documentBleedUniformSize = true;
  d.marginPreferences.left = d.marginPreferences.top = d.marginPreferences.right = d.marginPreferences.bottom = 0;
  d.marginPreferences.columnCount = 1;
  d.marginPreferences.columnGutter = 0;
  d.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  d.zeroPoint = [0, 0];
  if (d.masterSpreads[0].pages.length > 1) d.masterSpreads[0].pages[1].remove();
  var mp = d.masterSpreads[0].pages[0].marginPreferences;
  mp.left = mp.top = mp.right = mp.bottom = 0;
  mp.columnCount = 1;
  mp.columnGutter = 0;
  d.documentPreferences.pageWidth = w;
  d.documentPreferences.pageHeight = h;
  return d;
}

function removeDuplicates(array) {
  var unique = [];
  for (var i = 0; i < array.length; i++) {
    var found = false;
    for (var j = 0; j < unique.length; j++) {
      if (array[i] === unique[j]) { found = true; break; }
    }
    if (!found) unique.push(array[i]);
  }
  return unique;
}

function find_files(dir, mask_array) {
  var arr = [];
  for (var i = 0; i < mask_array.length; i++) arr = arr.concat(Folder(dir).getFiles("*" + mask_array[i]));
  return arr;
}

function padWithZeros(number, length) {
  var str = number.toString();
  var len = length.toString().length;
  while (str.length < len) str = "0" + str;
  return str;
}

function writeReportsErrors(arrErrors, pathofmaster) {
  var file = new File(pathofmaster + "/" + "_Errors_report" + dateAnZiOraMin + "_.txt");
  file.encoding = "UTF-8";
  file.open("w");
  for (var a = 0; a < arrErrors.length; a++) file.write(arrErrors[a] + "\r");
  file.close();
  alert("Some errors found, check Error report!");
}

function aliniere_zece_la_suta(myParent, mychild) {
  var pm = getPageMetrics();
  var w = pm.w, h = pm.h, w_t = pm.w_t, h_t = pm.h_t;
  var m_left = pm.m_left, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Y_HW_final = m_top + h - HW_h_final;
  var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;

  myParent.redefineScaling([1, 1]);
  mychild.redefineScaling([1, 1]);
  myParent.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left];
  app.documents[0].align(mychild, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
  myParent.fit(FitOptions.CENTER_CONTENT);

  var W_hw = myParent.geometricBounds[3] - myParent.geometricBounds[1];
  var W_hwai = mychild.geometricBounds[3] - mychild.geometricBounds[1];
  var H_hwai = mychild.geometricBounds[2] - mychild.geometricBounds[0];
  var anchor = AnchorPoint.CENTER_ANCHOR;
  var pw, matrix;

  if (w > h) {
    pw = (HW_h_final * 0.7) / H_hwai;
    matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
    mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }
  if (w < h) {
    myParent.redefineScaling([1, 1]);
    mychild.redefineScaling([1, 1]);
    pw = (W_hw * 0.8) / W_hwai;
    matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
    mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }

  W_hwai = mychild.geometricBounds[3] - mychild.geometricBounds[1];
  H_hwai = mychild.geometricBounds[2] - mychild.geometricBounds[0];

  if (W_hwai > w - w * 0.1) {
    myParent.redefineScaling([1, 1]);
    mychild.redefineScaling([1, 1]);
    pw = (W_hw * 0.8) / W_hwai;
    matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
    mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  }

  mychild.fit(FitOptions.CENTER_CONTENT);
  mychild.fit(FitOptions.frameToContent);
  myParent.fit(FitOptions.CENTER_CONTENT);
  myParent.geometricBounds = [Y_HW_final_latotal, -14.17322835, h_t + 14.17322835, w_t + 14.17322835];
}

function aliniere_cinci_la_suta(myParent, mychild) {
  var pm = getPageMetrics();
  var h = pm.h, w_t = pm.w_t, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Hdt_final = (h - HW_h_final) * DtiSize;
  var Y_DT_final = m_top + h - (HW_h_final + Hdt_final);

  myParent.redefineScaling([1, 1]);
  mychild.redefineScaling([1, 1]);
  myParent.geometricBounds = [Y_DT_final, -14.174, Y_DT_final + Hdt_final, w_t + 14.174];
  myParent.fit(FitOptions.CENTER_CONTENT);
  mychild.fit(FitOptions.frameToContent);
}

function aliniere_linia(myLine) {
  var pm = getPageMetrics();
  var h = pm.h, w_t = pm.w_t, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Hdt_final = (h - HW_h_final) * DtiSize;
  var lineY = m_top + h - (HW_h_final + Hdt_final) + Hdt_final;

  myLine.redefineScaling([1, 1]);
  myLine.geometricBounds = [lineY, -14.174, lineY, w_t + 14.174];
  myLine.strokeWeight = 1;
}
