# WkHtmlToPdf.Wrapper
C# wrapper for the WkHtmlToPdf executable - this library wraps the PDF generator from https://wkhtmltopdf.org/

# Usage
```
PdfGenerator pdfGenerator = new PdfGenerator();
PdfGenerator.Settings settings = new PdfGenerator.Settings 
{
	PageOrientation = PdfGenerator.PageOrientation.Landscape 
};
byte[] pdfBytes = pdfGenerator.GeneratePdf("https://www.google.com/", settings);
```
