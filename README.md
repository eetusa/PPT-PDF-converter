# PPT-PDF-converter
A tool to split PDF/PPT files in to image per page and parse through their text data, creating a JSON that links text and pages. Takes the desired target folder as an argument.

Definitely should have used a package to construct JSON and not parse it manually.

Used in [Slide Browser](https://github.com/eetusa/SlideBrowser), which is a tool to view PDF/PPT and sort them through text content.

## Other information
Uses Apache POI-HSLF for splitting the PPT to images and parsing thourgh the text, Apache PDFBox to split PDF to images and XpdfReader pdftotext to parse through PDF text.

## Known bugs
Spaces on folder paths cause it not to work, also " marks (quotes) in text content. 

