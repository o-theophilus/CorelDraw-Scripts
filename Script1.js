lo = "C:\\Users\\bo2\\Desktop";
// host.ActiveDocument.ExportBitmap(
// 	lo,
// 	host.cdrPNG,
// 	host.cdrCurrentPage,
// 	host.cdrRGBColorImage,
// 	100,
// 	100,
// 	300,
// 	300,
// 	host.cdrNormalAntiAliasing,
// 	false,
// 	true,
// 	true,
// 	false,
// 	host.cdrCompressionNone,
// 	undefined,
// 	undefined
// ).Finish();
// alert(lo)

host.ActiveDocument.Export(
	lo,
	host.cdrJPEG,
	host.cdrCurrentPage,
	{
		ImageType: cdrRGBColorImage,
		SizeX: 100,
		SizeY: 100
	},
	undefined,
	undefined
);
