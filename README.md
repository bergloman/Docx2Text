Docx2Text
===========================

This Windows console application (requires .NET framework 4.5 or newer to run) 
extracts given MS Word file (DOCX format) to plain text.

Usage:

	DocxText.exe <input file> <output file>
	DocxText.exe -d <input directory> <output directory>

This utility can be used to create backup of MS Word file(s) into `GIT` or some other source control provider
for backup purposes without saving binary files.

Uses `DocumentFormat.OpenXml` library internally. Use nuget to get the library from Nuget.org.
