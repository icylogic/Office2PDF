Msgbox, 请关闭所有需要转换的文档
FileSelectFolder, officeFolder,,, 选择原始文档所在目录
FileSelectFolder, pdfFolder,,, 选择PDF输出目录
ComObjError(false) 

wordApp:=comobjcreate("word.application")
Loop, %officeFolder%\*.doc?, , 1
{
    StringReplace, outFolder, A_LoopFileDir, %officeFolder%, , All
    SplitPath, A_LoopFileName,,,, outNameNoExt
    FileCreateDir, %pdfFolder%%outFolder%
    outputFile:= pdfFolder . outFolder . "\" . OutNameNoExt
    doc2pdf(wordApp, A_LoopFileLongPath, outputFile)
}
wordApp.quit(0)

pptApp:=comobjcreate("Powerpoint.Application")
Loop, %officeFolder%\*.ppt?, , 1
{
    StringReplace, outFolder, A_LoopFileDir, %officeFolder%, , All
    SplitPath, A_LoopFileName,,,, outNameNoExt
    FileCreateDir, %pdfFolder%%outFolder%
    outputFile:= pdfFolder . outFolder . "\" . OutNameNoExt
    ppt2pdf(pptApp, A_LoopFileLongPath, outputFile)
}
pptApp.quit

Loop, %officeFolder%\*.pdf, , 1
{
    StringReplace, outFolder, A_LoopFileDir, %officeFolder%, , All
    FileCreateDir, %pdfFolder%%outFolder%
    outFolder:= pdfFolder . outFolder . "\"
    FileCopy, %A_LoopFileLongPath%, %outFolder%
}

Run, %pdfFolder%

doc2pdf(wordApp, docFile, pdfFile){
    wdFormatPDF:=17
    wordApp.documents.Open(docFile)
    wordApp.documents(1).ExportAsFixedFormat(pdfFile, wdFormatPDF)
    wordApp.visible:=0
    word.documents(1).Close
}

ppt2pdf(pptApp, pptFile, pdfFile){
    ppSaveAsPDF:=32
    pptApp.Presentations.Open(pptFile,,,0)
    pptApp.Presentations(1).SaveAs(pdfFile, ppSaveAsPDF)
    pptApp.Presentations(1).Close
}