// Version-: 23-08-2017

//////////////////////////////////////contents.xml format/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//<office:spreadsheet>																			//
// 																					//
// <table:table table:name="name" table:style-name="ta1">								                    			        //
// 																					//
// 	<table:table-row table:number-rows-repeated="2" table:style-name="ro1">												//
//																					//
// 	<table:table-cell table:formula="of:=3*[.B2]"  table:number-columns-repeated="2" table:style-name="ce1" office:value-type="string" calcext:value-type="string">	//
// 	<text:p>SrNo<text:span text:style-name="T1">gj</text:span></text:p>												//
// 	</table:table-cell>																		//
//																            			        //
// </table:table-row>																			//
//																					//
// </office:spreadsheet>																		//
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

package ods

import (
	"archive/zip"
	"io/ioutil"
	"strconv"
)

type Cell struct {
	Type    string //Type float,string ...    ( office:value-type= )
	Value   string //Value                    ( office:value= )
	Formula string //formula 		  (table:formula= )
	Text    string //Text

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

			cellformulapara := "table:formula=\""
			cellformulaparalen := len(cellformulapara)
			cellformulaparaflag := false

			textspanstart := "<text:span"
			textspanstartlen := len(textspanstart)
			textspanstartflag := false

			textspanend := "</text:span>"
			textspanendlen := len(textspanend)

			para_start := "<text:p>"
			parastartlen := len(para_start)
			para_end := "</text:p>"
			paraendlen := len(para_end)
			para_started := false
			parastarti := -1
			xmlline := string(xmlfile)
			over := false
			tablename := ""

			rowparaflag := false
			cellparaflag := false
			tableparaflag := false
			blankcells:=0
			blankrows:=0
			blankcell:=Cell{"", "", "", ""}
			blankrow:=[]Cell{}
			blankrow=append(blankrow,blankcell)
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
								if firstcell<lastcell {
									for i := 0; i < blankrows; i++ {
										RowContents = append(RowContents, Row{blankrow})
										Rowno = Rowno + 1
									}
								  
								  
									for i := 0; i < row_repeatvalue; i++ {
										RowContents = append(RowContents, Row{CellContents[firstcell:lastcell]})
										Rowno = Rowno + 1
									}
									blankrows=0
								} else {
								      blankrows=blankrows+row_repeatvalue
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
								if (xmlline[i:i+cellendlen] == cell_end || (cellparaflag && xmlline[i:i+2] == "/>")) && !para_started {
									cell_started = false
									cellparaflag = false
									if len(celltext)>0 {
									      for i := 0; i < blankcells; i++ {
										      CellContents = append(CellContents, Cell{"", "", "", ""})
										      Cellno = Cellno + 1
									      }
									      for i := 0; i < column_repeatvalue; i++ {
										      CellContents = append(CellContents, Cell{celltype, cellvalue, cellformula, celltext})
										      Cellno = Cellno + 1
									      }
									      blankcells=0
									} else {
									    blankcells=blankcells+column_repeatvalue  
									}
									column_repeatvalue = 1
								}
								if para_started {
									if xmlline[i:i+paraendlen] == para_end || xmlline[i:i+2] == "/>" {
										para_started = false
										celltext = celltext + xmlline[parastarti:i]
									}
									///exclude text:span
									if xmlline[i:i+textspanstartlen] == textspanstart {
										textspanstartflag = true
										celltext = celltext + xmlline[parastarti:i]
									}
									if xmlline[i:i+1] == ">" && textspanstartflag {
										textspanstartflag = false
										parastarti = i + 1
									}
									if xmlline[i:i+textspanendlen] == textspanend {
										celltext = celltext + xmlline[parastarti:i]
										parastarti = i + textspanendlen
									}

								} else {
									if xmlline[i:i+parastartlen] == para_start {
										para_started = true
										parastarti = i + parastartlen
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
								/////////////////----cell parameters---////////////////////////

							} else {
								if xmlline[i:i+cellstartlen] == cell_start {
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
							if xmlline[i:i+rowstartlen] == row_start {
								row_started = true
								rowparaflag = true
								row_repeatvalue = 1
								firstcell = Cellno
								blankcells=0
							}
						}

					} else {
						if xmlline[i:i+tablestartlen] == table_start {
							table_started = true
							tableparaflag = true
							blankrows=0
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
