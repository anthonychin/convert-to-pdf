var fso = new ActiveXObject("Scripting.FileSystemObject");
var docPath = WScript.Arguments(0);
WScript.Echo("name is " + docPath);

docPath = fso.GetAbsolutePathName(docPath);

var pdfPath = docPath.replace(/\.ppt[^.]*$/, ".pdf");
var objWord = null;

try
{
    WScript.Echo("Saving '" + docPath + "' as '" + pdfPath + "'...");

    objWord = new ActiveXObject("PowerPoint.Application");
    //objWord = new PowerPoint.Application;
    WScript.Echo("Pass");
    objWord.Visible = true;

    var objDoc = objWord.Presentations.Open(docPath);
WScript.Echo("objDoc is "+objDoc);
    var ppSaveAsPDF = 32;
    objDoc.SaveAs(pdfPath, ppSaveAsPDF);
    //objDoc.Close();

// objDoc.ExportAsFixedFormat(pdfPath,
//     PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
//     PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
//     MsoTriState.msoFalse,PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
//     PowerPoint.PpPrintOutputType.ppPrintOutputSlides,MsoTriState.msoFalse, null,
//     PowerPoint.PpPrintRangeType.ppPrintAll, string.Empty,true, true, true,
//     true, false, unknownType);

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
