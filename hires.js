// Recorded 1/26/2022
let pdf = host.ActiveDocument.PDFSettings;
pdf.PublishRange = 0; // CdrPDFVBA.pdfWholeDocument
pdf.PageRange = "";
pdf.Author = "";
pdf.Subject = "";
pdf.Keywords = "";
pdf.BitmapCompression = 2; // CdrPDFVBA.pdfJPEG
pdf.JPEGQualityFactor = 10;
pdf.TextAsCurves = false;
pdf.EmbedFonts = true;
pdf.EmbedBaseFonts = true;
pdf.TrueTypeToType1 = true;
pdf.SubsetFonts = true;
pdf.SubsetPct = 80;
pdf.CompressText = true;
pdf.Encoding = 1; // CdrPDFVBA.pdfBinary
pdf.DownsampleColor = true;
pdf.DownsampleGray = true;
pdf.DownsampleMono = true;
pdf.ColorResolution = 200;
pdf.MonoResolution = 600;
pdf.GrayResolution = 200;
pdf.Hyperlinks = true;
pdf.Bookmarks = true;
pdf.Thumbnails = false;
pdf.Startup = 0; // CdrPDFVBA.pdfPageOnly
pdf.ComplexFillsAsBitmaps = false;
pdf.Overprints = true;
pdf.Halftones = false;
pdf.MaintainOPILinks = false;
pdf.FountainSteps = 256;
pdf.EPSAs = 0; // CdrPDFVBA.pdfPostscript
pdf.pdfVersion = 0; // CdrPDFVBA.pdfVersion12
pdf.IncludeBleed = false;
pdf.Bleed = 31750;
pdf.Linearize = false;
pdf.CropMarks = false;
pdf.RegistrationMarks = false;
pdf.DensitometerScales = false;
pdf.FileInformation = false;
pdf.ColorMode = 3; // CdrPDFVBA.pdfNative
pdf.UseColorProfile = true;
pdf.ColorProfile = 1; // CdrPDFVBA.pdfSeparationProfile
pdf.EmbedFilename = "";
pdf.EmbedFile = false;
pdf.JP2QualityFactor = 10;
pdf.TextExportMode = 0; // CdrPDFVBA.pdfTextAsUnicode
pdf.PrintPermissions = 0; // CdrPDFVBA.pdfPrintPermissionNone
pdf.EditPermissions = 0; // CdrPDFVBA.pdfEditPermissionNone
pdf.ContentCopyingAllowed = false;
pdf.OpenPassword = "";
pdf.PermissionPassword = "";
pdf.EncryptType = 0; // CdrPDFVBA.pdfEncryptTypeNone
pdf.OutputSpotColorsAs = 0; // CdrPDFVBA.pdfSpotAsSpot
pdf.OverprintBlackLimit = 95;
pdf.ProtectedTextAsCurves = true;

let SaveOptions = host.CreateStructSaveAsOptions();
SaveOptions.EmbedVBAProject = false;
SaveOptions.Filter = cdrCDR;
SaveOptions.IncludeCMXData = false;
SaveOptions.Range = cdrAllPages;
SaveOptions.EmbedICCProfile = true;
SaveOptions.Version = cdrVersion15;
SaveOptions.KeepAppearance = true;

let _name = host.ActiveDocument.FullFileName;
host.ActivePage.FindShapes(undefined, cdrTextShape, true).ConvertToCurves();
host.ActiveDocument.PublishToPDF(`${_name.substring(0, _name.length - 3)}pdf`);
host.ActiveDocument.SaveAs(_name, SaveOptions);
host.ActiveDocument.Close();
