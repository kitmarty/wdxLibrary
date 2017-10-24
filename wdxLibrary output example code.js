/* wdxLibrary Example.arx
 * this report demonstrates features of wdxLibrary.js
 * Copyright (c) 2017 Nikita Martyanov
 * https://github.com/kitmarty/wdxLibrary
 */

var test = new wdxWord()                                            //creates new output document;

//text styles examples
test.outSection()                                                   //outputs section without special settings
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER})
test.outText("Welcome to wdxLibrary! Enjoy it!",{RUN_BOLD:true})
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER,PAR_SPACE_AFTER:0})
test.outText("wdxLibrary is a Javascript library that works over OpenXML4J and ooxml-schemas")
test.outBreak()
test.outText("and allows to create .docx files using their functionality.")
test.outBreak()
test.outText("It's designed to get report documents with special formatting from ARIS Platform products.")
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER,PAR_SPACE_AFTER:0})
test.outText("This document generated using wdxLibrary",{RUN_FONTSIZE:20,RUN_FONT:"Courier New"})
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER,PAR_SPACE_AFTER:0})
test.outText("Copyright (c) 2017 Nikita Martyanov",{RUN_FONTSIZE:20,RUN_FONT:"Courier New"})
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER,PAR_SPACE_AFTER:0})
test.outText("https://github.com/kitmarty/wdxLibrary",{RUN_FONTSIZE:20,RUN_FONT:"Courier New"})
test.outLine()                                                      //outputs paragraph without special settings
test.outLine()                                                      //outputs paragraph without special settings
test.outText("Run without special settings")                        //outputs run without special settings
test.outLine()
test.outText("Текст кириллицей")                                    //outputs cyrillic text
test.outLine()
test.outText("Italic",{RUN_ITALIC:true})                            //outputs italic
test.outLine()
test.outText("Bold",{RUN_BOLD:true})                                //outputs bold
test.outLine()
test.outText("Italic&Bold",{RUN_ITALIC:true,RUN_BOLD:true})         //outputs bold and italic simultaneously
test.outLine()
test.outText("Single underline ",{RUN_UNDERLINE:wdxConstants.UL_SINGLE})
test.outText("Wavy heavy underline",{RUN_UNDERLINE:wdxConstants.UL_WAVY_HEAVY})
test.outText("superscript",{RUN_SCRIPT:wdxConstants.SCR_SUPERSCRIPT})
test.outLine()
test.outText("All letters are capitals ",{RUN_CAPS:true})
test.outText("subscript",{RUN_SCRIPT:wdxConstants.SCR_SUBSCRIPT})
test.outLine()
test.outText("All letters are small capitals",{RUN_SMALL_CAPS:true})
test.outLine()
test.outText("Striked text",{RUN_STRIKE:true})
test.outLine()
test.outText("Highlighted text",{RUN_HIGHLIGHT:wdxConstants.HL_CYAN})
test.outLine()
test.outText("Additional info about properties and constants you can find in wdxLibrary.js.")
test.outLine()
test.outText("Also you can read ECMA standarts and explore ooxml-schemas-x.x.jar for more functionality. It's easy to add to this library special formatting you need.")
test.outLine()
test.outText("Below you can find examples of interesting formatting such as footnotes, endnotes, embedded files etc.")

//page orientation, page size and page columns example
test.outSection({PAGE_COL_COUNT:3,PAGE_ORIENT:wdxConstants.PAGE_ORIENT_LANDSCAPE,PAGE_SIZE:wdxConstants.PAGE_SIZE_A4})
test.outLine()
test.outText("This section has 3 columns and landscape orientation.")
test.outBreak({BREAK_TYPE:wdxConstants.BREAK_COLUMN})
test.outText("Text in the second column")
test.outBreak({BREAK_TYPE:wdxConstants.BREAK_COLUMN})
test.outText("Text in the third column")

//footnotes and endnotes
test.outSection({SECTION_ENDNOTE:{EDN_NUMFMT:wdxConstants.NUMFMT_UPPER_LETTER}})
test.outLine()
test.outText("Text in paragraph with a footnote")
test.outFootnote({RUN_BOLD:true},{},{RUN_COLOR:"AA4455"})           //reference in text - bold, color of reference in footnote - AA4455
test.outText("footnote description without special formatting")
test.endFootnote()
test.outText(".")
test.outLine()
test.outText("Text in paragraph with an endnote")
test.outEndnote({RUN_BOLD:true},{},{RUN_COLOR:"CC8855"})           //reference in text - bold, color of reference in endnote - CC8855
test.outText("endnote description without special formatting")
test.endEndnote()
test.outText(". Description of this endnote you can find in the end of the document.")
test.outLine()
test.outText("Endnotes numbering format is upper letter. It's defined in the section properties")

//headers and footers
test.outSection()
test.outLine()
test.outText("This section has a header.")
test.outHeader()
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_CENTER})
test.outText("Bold text in header",{RUN_BOLD:true})
test.endHeader()
test.outLine()
test.outText("This text outputs after the header has been put in the file.")
test.outFooter()
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_RIGHT})
test.outField(wdxConstants.FLD_PAGE,{RUN_UNDERLINE:wdxConstants.UL_SINGLE})
test.endFooter()
test.outLine()
test.outText("This section also has a footer which contents page field with special formatting.")

//tables and bookmarks
test.outSection({SECTION_TYPE:wdxConstants.SM_NEXT_PAGE})//check wdxLibrary.js for more options of sections
test.outHeader()//if you don't want to have header from previos page call empty header. footer is same
test.endHeader()
test.outFooter()
test.endFooter()
//when you create table you can define border style before
//you can define every property of the table you find in ooxml-schemas
//it's just simple example
var brdStyle = {
    BORDER_COLOR:"243A84",
    BORDER_STYLE:wdxConstants.BORDER_DOT_DASH
}
var cellBrdStyle = {
    TBL_CELL_BORDER_TOP:brdStyle,
    TBL_CELL_BORDER_BOTTOM:brdStyle,
    TBL_CELL_BORDER_LEFT:brdStyle,
    TBL_CELL_BORDER_RIGHT:brdStyle
}
var cellStyle = {
    TBL_CELL_BORDERS:cellBrdStyle,
}
var bm1 = test.addBookmarkStart("Bookmark_name")//adding bookmark for example of internal hyperlink in the next section
test.outLine()
test.outText("This section demonstrates tables.")
test.addBookmarkEnd()
test.outLine()
test.outText("This section doesn't have header and footer.")
test.outTable({TBL_LAYOUT:wdxConstants.TBL_LAYOUT_AUTOFIT,TBL_WIDTH:{TBL_S_TYPE:wdxConstants.TBLW_PCT,TBL_S_WIDTH:5000}})
test.outRow()
test.outCell(cellStyle)
test.outLine()
test.outText("Colored text in cell",{RUN_COLOR:"AB1321"})
test.outCell(cellStyle).setStyle({TBL_CELL_SHADING:{TBL_SHADING_FILL:"AAAAAA"},TBL_CELL_WIDTH:{TBL_S_TYPE:wdxConstants.TBLW_PCT,TBL_S_WIDTH:2500}})
test.outLine()
test.outText("This table has two cells. Table is 100% width of the page, and cells are 50% width of the table.")
test.outBreak()
test.outText("This cell has #AAAAAA fill color.")
test.endTable()

//hyperlinks and numbering (lists)
test.outSection()
test.outLine()
test.outText("In the previous section I've added a bookmark to check this function.")
test.outLine()
test.outHyperlink("Hyperlink to the bookmark of the previous section",String(bm1.item.getName()))
test.outLine()
test.outText("Next hyperlink is external. Project page will be opened in default browser.")
test.outLine()
test.outHyperlink("https://github.com/kitmarty/wdxLibrary/",null,"https://github.com/kitmarty/wdxLibrary/")

//I don't like the numbering solution, but it works. Maybe I'll do it better if I have more spare time
var num = test.createNum()
num.setStyle({NUM_LVL:0})//0 - just to call case in switch block
num.setStyle({NUM_FMT:wdxConstants.NUMFMT_UPPER_ROMAN,NUM_ILVL:0,NUM_LVL_START:1,NUM_LVL_TEXT:"%1.",NUM_PSTYLE:{PAR_IND_LEFT:720,PAR_HANGING:360,PAR_CONT_SPACING:true}})
num.setStyle({NUM_LVL:0})
num.setStyle({NUM_FMT:wdxConstants.NUMFMT_DECIMAL,NUM_ILVL:1,NUM_LVL_START:1,NUM_LVL_TEXT:"%2)",NUM_PSTYLE:{PAR_IND_LEFT:1080,PAR_HANGING:360,PAR_CONT_SPACING:true}})
num.setStyle({NUM_LVL:0})
num.setStyle({NUM_FMT:wdxConstants.NUMFMT_BULLET,NUM_ILVL:2,NUM_LVL_START:1,NUM_LVL_TEXT:"",NUM_PSTYLE:{PAR_IND_LEFT:1440,PAR_HANGING:360,PAR_CONT_SPACING:true}})

test.outLine()
test.outText("There is an example of multilevel numbered (and bulleted) list below.")
test.outLine({PAR_LIST:0,PAR_LIST_ID:1})
test.outText("First level of the list. Upper Roman numbering format.")
test.outLine({PAR_LIST:1,PAR_LIST_ID:1})
test.outText("Second level of the list (decimal numfmt)")
test.outLine({PAR_LIST:1,PAR_LIST_ID:1})
test.outText("Second level of the list")
test.outLine({PAR_LIST:2,PAR_LIST_ID:1})
test.outText("Third level of the list (bullets)")
test.outLine({PAR_LIST:2,PAR_LIST_ID:1})
test.outText("Third level of the list (bullets)")
test.outLine({PAR_LIST:0,PAR_LIST_ID:1})
test.outText("First level of the list")

//pictures and embedded files
test.outSection()
test.outLine()
test.outText("There is an embedded file below (double click). ")
test.outText("Thumbnail generates on the fly using Java libraries. ")
test.outText("You can change it for ordinary MS Word icon but ")
test.outText("I don't know anything about copyrights that's why I've done this weird solution.")
test.outLine()
test.outEmbeddedFile(Context.getFile("wdxLibrary_empty.docx",Constants.LOCATION_COMMON_FILES),wdxConstants.FILE_TYPE_DOCX)
test.outLine()
test.outText("In this section you also can find examples of graphic output.")
test.outLine()
test.outText("The header of the section contents wdxLibrary logo and the page body contents picture of context model.")
test.outHeader()//if you don't want to have header from previos page call empty header. footer is same
test.outLine({PAR_ALIGN:wdxConstants.ALIGN_RIGHT})
test.outGraphic(Context.getFile("wdxLibrary.png",Constants.LOCATION_COMMON_FILES),wdxConstants.FILE_TYPE_PNG,75)
test.endHeader()
test.outLine()
test.outGraphic(ArisData.getSelectedModels()[0].Graphic(false,false,1049),wdxConstants.FILE_TYPE_EMF)

//native styles
test.outSection()
test.outLine()
test.outText("This section has the same header as in the previous section.")
test.outLine()
test.outText("wdxLibrary also allows you to define 'native' styles of the document. There is a paragraph below where the 'native' style is applied.")
test.createStyle({STYLE_TYPE:wdxConstants.STYLE_TYPE_PAR,STYLE_NAME:"Custom style by wdxLibrary",STYLE_BASED_ON:"Base",STYLE_RUN:{RUN_BOLD:true,RUN_UNDERLINE:wdxConstants.UL_WAVY_HEAVY},STYLE_PAR:{PAR_IND_LEFT:500}})
test.outLine({PAR_PSTYLE:"Custom style by wdxLibrary"})
test.outText("Text with the style that calls 'Custom style by wdxLibrary'")

test.writeReport()//you can specify filename or leave it empty.
