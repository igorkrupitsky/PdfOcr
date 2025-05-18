# Convert All Files to Searchable PDFs

Originally posted here: 
<https://www.codeproject.com/Articles/1303061/Convert-All-Files-to-Searchable-PDFs/>

## Introduction
This program will convert office, text and image files to PDFs. To use this program, drag your file(s) or folders onto the script file. Files in sub-folders will be converted too.

## Using the Code
The VBS script uses MS Office to convert Excel, Word, Text and Power Point documents.

The VBS script uses free Tesseract library (by Google) to convert images to PDF.

The script will add (_out) prefix to each PDF file. The prefix can be changed in Line2. Here is a script that will move all PDF files with (_out) prefix to a folder with (_out) prefix.

I have been using this script for some time and decided to share it. I hope someone else will find this useful. If you want to merge all of these PDFs, you can use the PDF Merge application I created earlier.
