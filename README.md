![image](/wdxLibrary.png)

## Introduction
wdxLibrary is a JavaScript library for generating Office Open XML Document (.docx) reports from ARIS Platform products. The main difference between the wdxLibrary and standard output engine is ability to customize wdxLibrary and get any Office Open XML Document formatting and features you need.

wdxLibrary is designed for ARIS script developers who need more functions from output engine than standard ARIS Report provides.

Read https://github.com/kitmarty/wdxLibrary/wiki for more information.

## Functions
wdxLibrary provides next document and text formatting:
* Sections and their properties
*	Columns of the page
*	Special text formatting (any formatting you can find in ECMA standards)
*	Special paragraph formatting
*	Footnotes
*	Endnotes
*	Headers
*	Footers
*	Fields (page number, table of contents etc)
*	Tables, table rows, table cells and their properties
*	Bookmarks
*	Hyperlinks (internal and external)
*	Numbered and bulleted lists (including multilevel lists)
*	Embedded files
*	Graphic output
*	“Native” styles of document

## Installation
wdxLibrary is based on several Java libraries. Most of them are already in ARIS Server (LOCAL or Business Server), but you have to download one more library (ooxmls-schemas-1.x.jar). This library provides an access to low-level document formatting functions. It’s free and you can find it here:

https://mvnrepository.com/artifact/org.apache.poi/ooxml-schemas

**Important!** You have to know the version of poi.jar that your ARIS Server uses. You can unzip poi.jar (find it in \lib directory of ARIS Server) and read META-INF/MANIFEST.MF to know its version. Use the table below to determine the version of ooxml-schemas.jar you need:

Version of poi.jar|Version of ooxml-schemas.jar
------------------|----------------------------
POI 3.5 and 3.6   |ooxml-schemas-1.0.jar
POI 3.7 to POI 3.13|ooxml-schemas-1.1.jar
POI 3.14 and newer|ooxml-schemas-1.3.jar

Additional info about ooxml-schemas you can find here:

https://poi.apache.org/faq.html#faq-N10025

**Important!** If you choose wrong version of ooxml-schemas.jar it will cause errors in ARIS Report functions and even standard reports will not work correctly.

Download ooxml-schemas–1.x.jar and put it in the server /lib directory.

For ARIS 7.2: {%ARIS_Installation_Directory%}/LocalServer/lib.

For ARIS 9.8 (9.6): {%ARIS_Installation_Directory%}/LOCALSERVER/bin/work/work_abs_local/base/webapps/abs/WEB-INF/lib

Don’t forget to restart server after adding ooxml-schemas–1.x.jar.

If you want to use it on Business Server/ Design Server/ ?Connect Server?, you can try to find lib directory where all ARIS Server libraries  (their names start with y-*) and poi.jar (poi-ooxml.jar etc) are stored.

## How to use

Download the latest version of wdxLibrary.js from https://github.com/kitmarty/wdxLibrary . Put it in the common files of Reports using ARIS Architect (report module).

Create new report. Import wdxLibrary.js from common files.

**Important!** If you use ARIS 7.2 there is no script option to get DOCX output files. Nevertheless you will get DOCX output file even if you choose DOC output. In this case next report that has same output filename will always rewrite previous output file. There are two solutions: set names dynamically in code or change output filename in the start report dialog.

See _wdxLibrary Example.arx_ to understand how to create reports using wdxLibrary.js.

Read https://github.com/kitmarty/wdxLibrary/wiki for more information.
