// Version-: 15-09-2017

//////////////////////////////////////contents.xml format/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//<office:spreadsheet>                                                                          							                //
//                                                                                                           								//
// <table:table table:name="name" table:style-name="ta1">                                                                                                     		//
//                                                                                                          								//
//      <table:table-row table:number-rows-repeated="2" table:style-name="ro1">                              								//
//                                                                                                           								//
//      <table:table-cell table:formula="of:=3*[.B2]"  table:number-columns-repeated="2" table:style-name="ce1" office:value-type="string" calcext:value-type="string"> //
//      <text:p>SrNo<text:span text:style-name="T1">gj</text:span><text:s text:c="10"/><text:s text:c="10"/>gh</text:p>                                                 //
//      </table:table-cell>                                                                                  								//
//                                                                                                                                                  			//
// </table:table-row>                                                                                        								//
//                                                                                                           								//
// </office:spreadsheet>                                                                                     								//
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

package ods

import (
	"archive/zip"
	"io/ioutil"
	"strconv"
	"strings"
)

type Cell struct {
	Type      string //Type float,string ...    ( office:value-type= )
	Value     string //Value                    ( office:value= )
	DateValue string //DateValue                ( office:date-value= )
	Formula   string //formula                  (table:formula= )
	Text      string //Text

}

type Row struct {
	Cells []Cell
}

type Sheet struct {
	Name string
	Rows []Row
}

type Odsfile struct {
	Sheets []Sheet
}

func ReplaceHTMLSpecialEntities(input string) string {
	output := strings.Replace(input, "&amp;", "&", -1)
	output = strings.Replace(output, "&lt;", "<", -1)
	output = strings.Replace(output, "&gt;", ">", -1)
	output = strings.Replace(output, "&quot;", "\"", -1)
	output = strings.Replace(output, "&lsquo;", "‘", -1)
	output = strings.Replace(output, "&rsquo;", "’", -1)
	output = strings.Replace(output, "&tilde;", "~", -1)
	output = strings.Replace(output, "&ndash;", "–", -1)
	output = strings.Replace(output, "&mdash;", "—", -1)
	output = strings.Replace(output, "&apos;", "'", -1)

	return output
}

func detectstart(checkinword string, startword string, startwordlength int) bool {
	yes := false
	if checkinword[0:startwordlength] == startword {
		if checkinword[startwordlength:startwordlength+1] == "/" {
			yes = true
		}
		if checkinword[startwordlength:startwordlength+1] == " " {
			yes = true
		}
		if checkinword[startwordlength:startwordlength+1] == ">" {
			yes = true
		}
		if checkinword[startwordlength-1:startwordlength] == ">" {
			yes = true
		}
	}
	return yes
}
func ReadODSFile(odsfilename string) (Odsfile, error) {
	var odsfileContents Odsfile
	r, err := zip.OpenReader(odsfilename)
	if err != nil {
		return odsfileContents, err
	}
	defer r.Close()

	Rowno := 0
	Cellno := 0
	firstcell := 0
	lastcell := 0
	firstrow := 0
	lastrow := 0
	CellContents := []Cell{}
	var celltext string
	var cellvalue string
	var celldatevalue string
	var celltype string
	var cellformula string
	RowContents := []Row{}
	SheetContents := []Sheet{}

	for _, f := range r.File {
		if f.Name == "content.xml" {
			rc, fileerr1 := f.Open()
			if fileerr1 != nil {
				return odsfileContents, fileerr1
			}
			xmlfile, fileerr := ioutil.ReadAll(rc)
			if fileerr != nil {
				return odsfileContents, fileerr
			}
			odsend := "</office:spreadsheet>"
			odsendlength := len(odsend)

			table_start := "<table:table"
			tablestartlen := len(table_start)
			table_end := "</table:table>"
			tableendlen := len(table_end)
			table_started := false
			table_name := "table:name=\""
			table_namelen := len(table_name)
			table_nameflag := false

			row_start := "<table:table-row"
			rowstartlen := len(row_start)
			row_end := "</table:table-row>"
			rowendlen := len(row_end)
			row_started := false

			cell_start := "<table:table-cell"
			cellstartlen := len(cell_start)
			cell_end := "</table:table-cell>"
			cellendlen := len(cell_end)
			cell_started := false

			column_repeat := "table:number-columns-repeated=\""
			column_repeatlen := len(column_repeat)
			column_repeatvalue := 1
			column_repeatflag := false

			row_repeat := "table:number-rows-repeated=\""
			row_repeatlen := len(row_repeat)
			row_repeatvalue := 1
			row_repeatflag := false

			celltypepara := "office:value-type=\""
			celltypeparalen := len(celltypepara)
			celltypeparaflag := false

			cellvaluepara := "office:value=\""
			cellvalueparalen := len(cellvaluepara)
			cellvalueparaflag := false

			celldatevaluepara := "office:date-value=\""
			celldatevalueparalen := len(celldatevaluepara)
			celldatevalueparaflag := false

			cellformulapara := "table:formula=\""
			cellformulaparalen := len(cellformulapara)
			cellformulaparaflag := false

			textspanstart := "<text:span"
			textspanstartlen := len(textspanstart)
			textspanparaflag := false
			textspanstarted := false
			textspanend := "</text:span>"
			textspanendlen := len(textspanend)

			annotationstart := "<office:annotation"
			annotationstartlen := len(annotationstart)
			annotationend := "</office:annotation>"
			annotationendlen := len(annotationend)
			annotationstarted := false
			annotation_paraflag := false

			para_start := "<text:p"
			parastartlen := len(para_start)
			para_end := "</text:p>"
			paraendlen := len(para_end)
			para_started := false
			para_paraflag := false

			textspacestart := "<text:s"
			textspacestartlen := len(textspacestart)
			textspacestarted := false
			textspaceparaflag := false
			textspaceparapresent := false

			text_spacec := "text:c=\""
			text_spacecflag := false
			text_spaceclen := len(text_spacec)
			parastarti := -1
			xmlline := string(xmlfile)
			over := false
			tablename := ""

			rowparaflag := false
			cellparaflag := false
			tableparaflag := false
			blankcells := 0
			blankrows := 0
			blankcell := Cell{"", "", "", "", ""}
			blankrow := []Cell{}
			blankrow = append(blankrow, blankcell)
			for i := 0; i < len(xmlline); i++ {
				if !over {
					if table_started {
						if (xmlline[i:i+tableendlen] == table_end || (tableparaflag && xmlline[i:i+2] == "/>")) && !row_started {
							table_started = false
							lastrow = Rowno
							SheetContents = append(SheetContents, Sheet{tablename, RowContents[firstrow:lastrow]})
							tablename = ""

						}
						if xmlline[i:i+1] == ">" {
							tableparaflag = false
						}
						/////////////////table parameters////////////////////////
						if table_nameflag {
							if xmlline[i:i+1] == "\"" {
								if parastarti < i {
									table_nameflag = false
									tablename = xmlline[parastarti:i]
								}
							}
						} else {
							if xmlline[i:i+table_namelen] == table_name && tableparaflag {
								table_nameflag = true
								parastarti = i + table_namelen
							}
						}
						/////////////////---table parameters--////////////////////////

						if row_started {
							if (xmlline[i:i+rowendlen] == row_end || (rowparaflag && xmlline[i:i+2] == "/>")) && !cell_started {
								row_started = false
								lastcell = Cellno
								rowparaflag = false
								if firstcell < lastcell {
									for i := 0; i < blankrows; i++ {
										RowContents = append(RowContents, Row{blankrow})
										Rowno = Rowno + 1
									}

									for i := 0; i < row_repeatvalue; i++ {
										RowContents = append(RowContents, Row{CellContents[firstcell:lastcell]})
										Rowno = Rowno + 1
									}
									blankrows = 0
								} else {
									blankrows = blankrows + row_repeatvalue
								}
								row_repeatvalue = 1
							}

							if xmlline[i:i+1] == ">" {
								rowparaflag = false
							}
							/////////////////row parameters////////////////////////
							if row_repeatflag {
								if xmlline[i:i+1] == "\"" {
									if parastarti < i {
										row_repeatflag = false
										value, err := strconv.Atoi(xmlline[parastarti:i])
										if err != nil {
											return odsfileContents, err
										}
										row_repeatvalue = value
									}

								}
							} else {
								if xmlline[i:i+row_repeatlen] == row_repeat && rowparaflag {
									row_repeatflag = true
									parastarti = i + row_repeatlen
								}
							}
							/////////////////----row parameters---////////////////////////

							if cell_started {
								if xmlline[i:i+1] == ">" {
									cellparaflag = false
								}

								if (xmlline[i:i+cellendlen] == cell_end || (cellparaflag && xmlline[i:i+2] == "/>")) && !para_started && !annotationstarted {
									cell_started = false
									cellparaflag = false
									if len(celltext) > 0 || len(cellvalue) > 0 || len(celldatevalue) > 0 || len(cellformula) > 0 {
										for i := 0; i < blankcells; i++ {
											CellContents = append(CellContents, Cell{"", "", "", "", ""})
											Cellno = Cellno + 1
										}
										for i := 0; i < column_repeatvalue; i++ {
											////////////////////--Replacing HTML SPECIAL ENTITIES--////////////////////////
											celltext = ReplaceHTMLSpecialEntities(celltext)
											cellvalue = ReplaceHTMLSpecialEntities(cellvalue)
											///////////////////////////////////////////////////////////////////////////////
											CellContents = append(CellContents, Cell{celltype, cellvalue, celldatevalue, cellformula, celltext})
											Cellno = Cellno + 1
										}
										blankcells = 0
									} else {
										blankcells = blankcells + column_repeatvalue
									}
									column_repeatvalue = 1
								}
								if annotationstarted {
									if xmlline[i:i+1] == ">" && annotation_paraflag {
										annotation_paraflag = false
									}
									if xmlline[i:i+annotationendlen] == annotationend || (annotation_paraflag && xmlline[i:i+2] == "/>") {
										annotationstarted = false
										annotation_paraflag = false
									}

								} else {
									if detectstart(xmlline[i:i+annotationstartlen+2], annotationstart, annotationstartlen) {
										annotationstarted = true
										annotation_paraflag = true
									}
								}

								if para_started {
									if xmlline[i:i+1] == ">" && para_paraflag {
										para_paraflag = false
										parastarti = i + 1
									}
									if (xmlline[i:i+paraendlen] == para_end || (para_paraflag && xmlline[i:i+2] == "/>")) && !textspanstarted && !textspacestarted {
										para_started = false
										if !para_paraflag {
											celltext = celltext + xmlline[parastarti:i]
											parastarti = i + paraendlen
										} else {
											parastarti = i + 2
										}
										para_paraflag = false
									}

									if textspanstarted {
										if !textspacestarted {
											if xmlline[i:i+1] == ">" {
												textspanparaflag = false
												parastarti = i + 1
											}
											if xmlline[i:i+textspanendlen] == textspanend || (textspanparaflag && xmlline[i:i+2] == "/>") {
												textspanstarted = false
												if !textspanparaflag {
													celltext = celltext + xmlline[parastarti:i]
													parastarti = i + textspanendlen
												} else {
													parastarti = i + 2
												}
												textspanparaflag = false

											}
										}
									} else {
										if detectstart(xmlline[i:i+textspanstartlen+2], textspanstart, textspanstartlen) {
											textspanparaflag = true
											textspanstarted = true
											celltext = celltext + xmlline[parastarti:i]
										}
									}
									if textspacestarted {
										if xmlline[i:i+1] == ">" {
											textspaceparaflag = false
											parastarti = i + 1
										}
										if textspaceparaflag && xmlline[i:i+2] == "/>" {
											textspacestarted = false
											parastarti = i + 2
											textspaceparaflag = false
											if !textspaceparapresent {
												celltext = celltext + " "
											}
										}
										/////////////////space parameters////////////////////////
										if text_spacecflag {
											if xmlline[i:i+1] == "\"" {
												if parastarti < i {
													text_spacecflag = false
													value, err := strconv.Atoi(xmlline[parastarti:i])
													if err != nil {
														return odsfileContents, err
													}
													for i := 0; i < value; i++ {
														celltext = celltext + " "
													}
													textspaceparapresent = true

												}
											}
										} else {
											if xmlline[i:i+text_spaceclen] == text_spacec && textspaceparaflag {
												text_spacecflag = true
												parastarti = i + text_spaceclen
											}
										}
										//////////////////////////////////////////////////////////
									} else {
										if detectstart(xmlline[i:i+textspacestartlen+2], textspacestart, textspacestartlen) {
											textspaceparaflag = true
											textspacestarted = true
											textspaceparapresent = false
											celltext = celltext + xmlline[parastarti:i]
										}
									}
								} else {
									if detectstart(xmlline[i:i+parastartlen+2], para_start, parastartlen) && !annotationstarted {
										para_started = true
										para_paraflag = true
										parastarti = i + parastartlen
										if len(celltext) > 0 {
											celltext = celltext + "\n"
										}
									}

								}

								/////////////////cell parameters////////////////////////
								if column_repeatflag {
									if xmlline[i:i+1] == "\"" {
										if parastarti < i {
											column_repeatflag = false
											value, err := strconv.Atoi(xmlline[parastarti:i])
											if err != nil {
												return odsfileContents, err
											}
											column_repeatvalue = value
										}
									}
								} else {
									if xmlline[i:i+column_repeatlen] == column_repeat && cellparaflag {
										column_repeatflag = true
										parastarti = i + column_repeatlen
									}
								}
								if celltypeparaflag {
									if xmlline[i:i+1] == "\"" {
										if parastarti < i {
											celltypeparaflag = false
											celltype = xmlline[parastarti:i]
										}
									}
								} else {
									if xmlline[i:i+celltypeparalen] == celltypepara && cellparaflag {
										celltypeparaflag = true
										parastarti = i + celltypeparalen
									}

								}
								if cellformulaparaflag {
									if xmlline[i:i+1] == "\"" {
										if parastarti < i {
											cellformulaparaflag = false
											cellformula = xmlline[parastarti:i]
										}

									}
								} else {
									if xmlline[i:i+cellformulaparalen] == cellformulapara && cellparaflag {
										cellformulaparaflag = true
										parastarti = i + cellformulaparalen

									}
								}

								if cellvalueparaflag {
									if xmlline[i:i+1] == "\"" {
										if parastarti < i {
											cellvalueparaflag = false
											cellvalue = xmlline[parastarti:i]
										}
									}
								} else {
									if xmlline[i:i+cellvalueparalen] == cellvaluepara && cellparaflag {
										cellvalueparaflag = true
										parastarti = i + cellvalueparalen

									}
								}

								if celldatevalueparaflag {
									if xmlline[i:i+1] == "\"" {
										if parastarti < i {
											celldatevalueparaflag = false
											celldatevalue = xmlline[parastarti:i]
										}
									}
								} else {
									if xmlline[i:i+celldatevalueparalen] == celldatevaluepara && cellparaflag {
										celldatevalueparaflag = true
										parastarti = i + celldatevalueparalen

									}
								}
								/////////////////----cell parameters---////////////////////////

							} else {
								if detectstart(xmlline[i:i+cellstartlen+2], cell_start, cellstartlen) {
									cell_started = true
									cellparaflag = true
									column_repeatvalue = 1
									celltext = ""
									celltype = ""
									cellvalue = ""
									cellformula = ""

								}
							}

						} else {
							if detectstart(xmlline[i:i+rowstartlen+2], row_start, rowstartlen) {
								row_started = true
								rowparaflag = true
								row_repeatvalue = 1
								firstcell = Cellno
								blankcells = 0
							}
						}

					} else {
						if detectstart(xmlline[i:i+tablestartlen+2], table_start, tablestartlen) {
							table_started = true
							tableparaflag = true
							blankrows = 0
							table_nameflag = false
							row_started = false
							firstrow = Rowno
							cell_started = false
							cellvalueparaflag = false
							celltypeparaflag = false
							column_repeatvalue = 1
							row_repeatvalue = 1
							column_repeatflag = false
							row_repeatflag = false
						}
						if xmlline[i:i+odsendlength] == odsend {
							over = true
						}
					}
				}

			}
		}
	}

	odsfileContents.Sheets = SheetContents
	return odsfileContents, err
}
