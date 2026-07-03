Option Explicit

Sub CombinedExportProjects()
    ' ==============================================================
    ' Exports a CorelDRAW project as:
    '   - PDF certificate (5-page projects only, exports page 5)
    '   - JPG per page
    '   - Flattened/curved CDR (CorelDRAW v15) archive copy
    ' ==============================================================

    Dim sfolder As String
    Dim i As Long, j As Integer
    Dim p As Page
    Dim expflt As ExportFilter
    Dim pathPDF As String, pathCDR As String, pathJPG As String
    Dim opt As New StructSaveAsOptions
    Dim sr As ShapeRange
    Dim grp1 As ShapeRange
    Dim s1 As Shape

    sfolder = CorelScriptTools.GetFolder("E:\Yaser Arafath\OneDrive - ICT ACADEMY\Email attachments\Projects")
    j = ActiveDocument.Pages.Count

    ' Handle 5-page project (Honeywell)
    If j = 5 Then
        ' Export last page to PDF
        ActiveDocument.Pages(5).Activate
        pathPDF = sfolder & "\" & ActivePage.Name & ".pdf"

        With ActiveDocument.PDFSettings
            .PublishRange = 1 ' CdrPDFVBA.pdfCurrentPage
            .PageRange = "5"
            .Author = "Yaser ARafath B"
            .Subject = "CSR Certificate"
            .Keywords = "CSR, ICT ACademy"
            .BitmapCompression = 2 ' CdrPDFVBA.pdfJPEG
            .JPEGQualityFactor = 2
            .TextAsCurves = True
            .EmbedFonts = True
            .EmbedBaseFonts = False
            .TrueTypeToType1 = True
            .SubsetFonts = True
            .SubsetPct = 80
            .CompressText = True
            .Encoding = 1 ' CdrPDFVBA.pdfBinary
            .DownsampleColor = True
            .DownsampleGray = True
            .DownsampleMono = True
            .ColorResolution = 300
            .MonoResolution = 120
            .GrayResolution = 300
            .Hyperlinks = True
            .Bookmarks = False
            .Thumbnails = False
            .Startup = 0 ' CdrPDFVBA.pdfPageOnly
            .ComplexFillsAsBitmaps = True
            .Overprints = False
            .Halftones = False
            .MaintainOPILinks = False
            .FountainSteps = 256
            .EPSAs = 1 ' CdrPDFVBA.pdfPreview
            .pdfVersion = 0 ' CdrPDFVBA.pdfVersion12
            .IncludeBleed = True
            .Bleed = 30000
            .Linearize = True
            .CropMarks = True
            .RegistrationMarks = False
            .DensitometerScales = False
            .FileInformation = False
            .ColorMode = 0 ' CdrPDFVBA.pdfRGB
            '.UseColorProfile = True
            .ColorProfile = 1 ' CdrPDFVBA.pdfSeparationProfile
            .EmbedFilename = ""
            .EmbedFile = False
            .JP2QualityFactor = 2
            .TextExportMode = 0 ' CdrPDFVBA.pdfTextAsUnicode
            .PrintPermissions = 0 ' CdrPDFVBA.pdfPrintPermissionNone
            .EditPermissions = 0 ' CdrPDFVBA.pdfEditPermissionNone
            .ContentCopyingAllowed = False
            .OpenPassword = ""
            .PermissionPassword = ""
            .EncryptType = 2 ' CdrPDFVBA.pdfEncryptTypeAES
            .OutputSpotColorsAs = 0 ' CdrPDFVBA.pdfSpotAsSpot
            .OverprintBlackLimit = 95
            .ProtectedTextAsCurves = True
            .UsePageBoundingBox = False
        End With

        ActiveDocument.PublishToPDF pathPDF
        ActiveDocument.Pages(5).Delete
        MsgBox "Certificate Exported Successfully", vbOKOnly
        Application.Wait (Now + TimeValue("00:00:03"))
    ElseIf j > 5 Then
        MsgBox "Please open a 4-page or 5-page CorelDraw file.", vbOKOnly
        Exit Sub
    End If

    ' Export remaining pages to JPG
    ' NOTE: uses p.Shapes.All (not ActiveLayer.Shapes.All) so shapes on
    ' every layer of the page are included in the export, not just the
    ' active layer.
    For Each p In ActiveDocument.Pages
        p.Activate
        p.Shapes.All.CreateSelection

        pathJPG = sfolder & "\" & ActivePage.Name & ".jpg"

        Set expflt = ActiveDocument.ExportBitmap(pathJPG, cdrJPEG, cdrSelection, cdrRGBColorImage, 0, 0, 360, 360, cdrNormalAntiAliasing, False, False, True, False, cdrCompressionJPEG)
        With expflt
            .Smoothing = 50
            .Compression = 1
            .Finish
        End With
    Next p

    ' Convert text to curves on every page (protects fonts/editability)
    For i = 1 To ActiveDocument.Pages.Count
        ActiveDocument.Pages(i).Activate
        ActivePage.Shapes.All.CreateSelection
        Set grp1 = ActiveSelection.UngroupAllEx
        Set sr = ActivePage.Shapes.All
        If sr.Count > 0 Then
            ActivePage.Shapes.FindShapes(Query:="@type = 'text:artistic'").ConvertToCurves
            ActivePage.Shapes.FindShapes(Query:="@type = 'text:paragraph'").ConvertToCurves
        End If
        ActivePage.Shapes.All.CreateSelection
        Set s1 = ActiveSelection.Group
    Next i

    ' Save as CDR Version 15 (matches StructSaveAsOptions.Version below)
    pathCDR = sfolder & "\" & ActiveDocument.Name

    With opt
        .EmbedICCProfile = False
        .EmbedVBAProject = False
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Overwrite = True
        .Range = cdrAllPages
        .Version = cdrVersion15
        .KeepAppearance = True
    End With

    ActiveDocument.SaveAs pathCDR, opt

    MsgBox "Export process completed successfully!", vbInformation

    'ActiveDocument.Close
End Sub
