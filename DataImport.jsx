#target illustrator; 

/*
 *Copyright Kim Wu  kimwu@live.hk
 */

var isDebug = false;
var charFrom, charTo;
var charStyle, charAttr;
var docRef = app.activeDocument;
var tfCoo = docRef.pageItems.getByName("CountryOfOrigin");
var cooContents = "";

function filterFiles(file) {
    while (file.alias) {
        file = file.resolve();
        if (file == null) { return false }
    }
    if (file.constructor.name == "Folder") return true
    var extn = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
        switch (extn) {
            case ".csv" : return true;
    default : return false
    }
}
// get Fonts by fonts name
function getFont(fontName){
        try {
            var tFont = app.textFonts.getByName(fontName);
        }
        catch(e){
            alert("Fonts \"" + fontName + "\"  not found.", "Worning")
            try{
                var tFont = app.textFonts.getByName("DINCond-Medium");
            }
            catch(e){
                var tFont = app.textFonts[0]
             };
         }
        return tFont;
};

function getCharStyle(charStyleName){
            try {
                    var charStyle = app.textFonts.getByName(charStyleName);
               }
             catch ( e){
                    var charStyle = docRef.characterStyles.add(charStyleName);
                }
            return charStyle;
};


var dfFolder=Folder.desktop;

if (File.fs == "Windows"){
    var   fileFillter =  "CSV File:*.csv";
    }
else if (file.fs == "Macintosh"){
    var fileFillter = filterFiles ;
    }

if (isDebug) {
    var dataFile = File("/P/jsx4ai/SourceData.csv");
    }
else{
    var dataFile = dfFolder.openDlg("Please select a CSV file", fileFillter, false);
}    

dataFile.open("r");

var ln = "", lnData;
var idx = 0;
var contents = [];

//Read data from csv file, split into object list. 
do {
    ln = dataFile.readln();
    //dataFile.current
    if (ln.replace(/(^\s*)|(\s*$)/g,"").length > 0 &&  ln[0] !=";" && ln[0] !="#"){
        lnData = ln.split(" || ");
        contents[contents.length] ={};
        for (index =0; index < lnData.length; index++) {
            tmpData = lnData[index].split(":");
            contents[contents.length - 1][tmpData[0].replace(/(^\s*)|(\s*$)/g,"")] = tmpData[1].replace(/(^\s*)|(\s*$)/g,"");
        }
    }
}
while (dataFile.eof == false);

var cooData = [];
var defaultStyle = [];

for (index = 0; index < contents.length; index++) {
    if (contents[index].dataType == "COOData"){
        cooData[cooData.length] = contents[index];
     }
    else if (contents[index].dataType == "DEFAULT"){
        defaultStyle[defaultStyle.length] = contents[index];
     }
}

if  (defaultStyle[0] == null){
    defaultStyle[0] = {lang:"ALL", charSize:5.5, fontName:"DINCond-Medium", text:"0123456789%\/"};
}


// Join all COO  contents and set the position. 
for (index = 0; index < cooData.length; index++) {
    cooData[index].position = cooContents.length;
    if ( index == (cooData.length - 1)) {
        cooContents += cooData[index].text;
        }
    else {
        cooContents += cooData[index].text + " \/ ";
        }
} 
// add coo contents to text frame. 
tfCoo.textRange.contents = cooContents;

// set  COO Char Sytle
for (index = 0; index < cooData.length; index++) {
        charStyle = getCharStyle(cooData[index].lang + " Style");
        charAttr = charStyle.characterAttributes;
        charAttr.size = cooData[index].charSize;
        charAttr.textFont = getFont(cooData[index].fontName);
        if (cooData[index].position + cooData[index].text.length + 3 < tfCoo.textRange.contents.length){
            charTo = cooData[index].position + cooData[index].text.length + 3;
            }
        else{
            charTo = tfCoo.textRange.contents.length;
            }
        for( i = cooData[index].position; i < charTo ; i++){
                if (tfCoo.textRange.characters[i] == "\/"){
                    // test
                }
                else {
                    charStyle.applyTo(tfCoo.textRange.characters[i]);
                    // tfCoo.textRange.characters[i].textFont = CharAttr.textFont
                  }
         }
} ;
