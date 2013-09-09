var fso = new ActiveXObject("Scripting.FileSystemObject");

// Get Files
var f = fso.GetFolder("."); //folder
var fc = new Enumerator(f.files); //folder content

var s = "";
var docPath = "";
var objWord = null;

// iterations through files
for (; !fc.atEnd(); fc.moveNext()){
    s = fc.item();
    s += "";
    docPath = s.replace(/^.*(\\|\/|\:|.ppt[^.]*|.js|.pdf|.jpg|.mp3|.png)/,'');
    WScript.Echo("docPath is" + docPath)

    if(docPath.length>5){
        docPath = fso.GetAbsolutePathName(docPath);

        var pdfPath = docPath.replace(/\.doc[^.]*$/, ".pdf");

        try
        {
            WScript.Echo("Saving '" + docPath + "' as '" + pdfPath + "'...");

            objWord = new ActiveXObject("Word.Application");
            objWord.Visible = false;

            var objDoc = objWord.Documents.Open(docPath);

            var wdFormatPdf = 17;
            objDoc.SaveAs(pdfPath, wdFormatPdf);
            objDoc.Close();

            WScript.Echo("Done.");
        }
        finally
        {
            if (objWord != null)
            {
                objWord.Quit();
            }
        }
    }
}