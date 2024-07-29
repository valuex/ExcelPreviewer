# ExcelPreviewer
Select the cell in MS Excel and preview the corresponding file.

# How to use it for previwing pdf files
1. Download **ExcelPreviewer.xlsm** and **XPDFViewer.zip**.
2. Unzip the **XPDFViewer.zip** and place it at somewhere you like, let's call it ` path of XPDFViewer.exe`.  
3. Open the **ExcelPreviewer.xlsm**, go To worksheet `Settings`, write  ` path of XPDFViewer.exe` into cell `B2`.
4. Click `Start` to open the **XPDFViewer.exe**
5. Go to the Worksheet `PDF` in **ExcelPreviewer.xlsm**, put the pdf file full name (with folder path) list in the first column.
6. Using key `UP` or `DOWN`  to preview the corresponding pdf file. That's it.

# How to use for previwing html files
1. Open the **ExcelPreviewer.xlsm**, go To Worksheet **HTML**, put the html file name list in the first column.
2. Open the **WebPagePreviewer.exe**
3. Go To Worksheet **HTML**, using key `UP` or `DOWN`  to preview the corresponding  file.

# Supported file types
1. for  Worksheet **PDF** and **XPDFViewer.exe**, one can only use it for pdf previewing.
2. for  Worksheet **HTML** and **WebPagePreviewer.exe**,  because webview2.dll is used for file previewing, so many types of file are supported, including:
   - webpage file: *.mhtml; *.htm
   - text file:
   - image file:
   - pdf file: yes, pdf is also supported, but with low speed, that Why I develop a specific appliccation for PDF files.
# Important Note
1. All these code are developed in and for 64bit application, so ensure that you get 64-bit Excel.
   
# Acknowledgement
1. **SumatraPDFControl** ： https://github.com/marcoscmonteiro/SumatraPDFControl
2. **AHK binding of Webview2.dll**: https://github.com/thqby/ahk2_lib/blob/master/WebView2/WebView2.ahk
