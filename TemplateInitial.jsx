#target illustrator; 

/*
 *Copyright Kim Wu  kimwu@live.hk
 * Create five textFrame
* Size  , Fabric Content ,   Country of Origin ,   Care Instruction,  Washing Symbols
*/

var docRef, layerRef;
var tfSize, tfFabricContent, tfCountryOfOrigin, tfCareInstruction, tfWashingSymbols;
var piSize, piFabricContent, piCountryOfOrigin, piCareInstruction, piWashingSymbols;
 
 if (app.documents.length > 0 ) {
     docRef = app.activeDocument;
     }
else {
    docRef = app.documents.add();
    }
 
 try {  
     layerRef = docRef.layers.getByName("ImportData");
     }
catch(e) { 
    layerRef = docRef.layers.add();
    layerRef.name = "ImportData";
    }  

try {
    tfSize = docRef.pageItems.getByName("Size");
    }
catch(e) {
    //alert(e.message,"Warning", true);
    piSize = docRef.pathItems.rectangle( -110, docRef.height - 20, docRef.width * 0.40, docRef.height * 0.25 )
    tfSize = docRef.textFrames.areaText(piSize);    
    tfSize.name = "Size";
    tfSize.contents = "Size";
    tfSize.position = [-110, docRef.height - 20];
    }

try {
    tfFabricContent = docRef.pageItems.getByName("FabricContent");
    }
catch(e) {
    piFabricContent = docRef.pathItems.rectangle( 10, docRef.height - 240, docRef.width * 0.40, docRef.height * 0.25 )
    tfFabricContent = docRef.textFrames.areaText(piFabricContent);    
    tfFabricContent.name = "FabricContent";
    tfFabricContent.contents = "Fabric Content";
    tfFabricContent.position = [10, docRef.height - 240];
    }

try {
    tfCountryOfOrigin = docRef.pageItems.getByName("CountryOfOrigin");
    }
catch(e) {
    //alert(e.message,"Warning", true);
    //tfCountryOfOrigin = layerRef.textFrames.add();
    piCountryOfOrigin = docRef.pathItems.rectangle( docRef.width * 0.52, docRef.height - 20, docRef.width * 0.40, docRef.height * 0.25 )
    tfCountryOfOrigin = docRef.textFrames.areaText(piCountryOfOrigin);    
    tfCountryOfOrigin.name = "CountryOfOrigin";
    tfCountryOfOrigin.contents = "Country Of Origin";
    tfCountryOfOrigin.position = [docRef.width * 0.52, docRef.height - 20];
    }

// Create a new character style
//var charStyle = docRef.characterStyles.add("BigRed");
// set character attributes
//var charAttr = charStyle.characterAttributes;
//charAttr.size = 40;
//charAttr.tracking = -50;
//charAttr.capitalization = FontCapsOption.ALLCAPS;
//var redColor = new RGBColor();
//~ redColor.red = 255;
//~ redColor.green = 0;
//~ redColor.blue = 0;
//~ charAttr.fillColor = redColor;
//~ var char2 = tfCountryOfOrigin.textRange.characters[2];
//~ char2.le
//~ tfCountryOfOrigin.textRange.characters[2].characterAttributes.horizontalScale = 300; 
//~ charStyle.applyTo(tfCountryOfOrigin.textRange.words[4]);

try {
    tfCareInstruction = docRef.pageItems.getByName("CareInstruction");
    }
catch(e) {
    piCareInstruction = docRef.pathItems.rectangle(docRef.width * 0.52, docRef.height - 240, docRef.width * 0.40, docRef.height * 0.25 )
    tfCareInstruction = docRef.textFrames.areaText(piCareInstruction);    
    tfCareInstruction.name = "CareInstruction";
    tfCareInstruction.contents = "Care Instruction";
    tfCareInstruction.position = [docRef.width * 0.52, docRef.height - 240];
    }

try {
    tfWashingSymbols = docRef.pageItems.getByName("WashingSymbols");
    }
catch(e) {
    //alert(e.message,"Warning", true);
    piWashingSymbols = docRef.pathItems.rectangle( 10, docRef.height - 480, docRef.width * 0.40, docRef.height * 0.25 )
    tfWashingSymbols = docRef.textFrames.areaText(piWashingSymbols);    
    tfWashingSymbols.name = "WashingSymbols";
    tfWashingSymbols.contents = "Washing Symbols";
    tfWashingSymbols.position = [10 , docRef.height - 480];
    }