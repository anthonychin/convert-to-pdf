var fso = new ActiveXObject("Scripting.FileSystemObject");

// Get Files
var f = fso.GetFolder(".");
var fc = new Enumerator(f.files);

var s = "";
var docPath = "";
var objWord = null;

// iterations through files
for (; !fc.atEnd(); fc.moveNext()){
    s = fc.item();
    s += "";
    docPath = s.replace(/^.*(\\|\/|\:|.doc[^.]|.js|.pdf|.jpg|.mp3|.png)/,'');

    WScript.Echo("docPath is" + docPath );

    // Check if the path is at least of length > 5
    if(docPath.length>5){
    docPath = fso.GetAbsolutePathName(docPath);

    var pdfPath = docPath.replace(/\.ppt[^.]*$/, ".pdf");

        try
        {
            WScript.Echo("Saving '" + docPath + "' as '" + pdfPath + "'...");

            objWord = new ActiveXObject("PowerPoint.Application");
            objWord.Visible = true;

            var objDoc = objWord.Presentations.Open(docPath);

            var ppSaveAsPDF = 32;
            objDoc.SaveAs(pdfPath, ppSaveAsPDF);

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