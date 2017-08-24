package main

import (
	"./ReadOdsFile"
	"os"
	"bufio"
	"fmt"
)


// writeLines writes the lines to the given file.
func writeLines(lines []string, path string) error {
  file, err := os.Create(path)
  if err != nil {
    return err
  }
  defer file.Close()

  w := bufio.NewWriter(file)
  for _, line := range lines {
    fmt.Fprintln(w, line+"\r")
  }
  return w.Flush()
}

func main() {
    Filecontent,eerr:=ods.ReadODSFile("test.ods")
    if eerr != nil {
		fmt.Printf("Read : %s\n", eerr)
		var yes string
		fmt.Scan(&yes)
		os.Exit(0)
    }
    for _,sheet := range Filecontent.Sheets {
	outputcontent:=[]string{}
	for _,row := range sheet.Rows {
	    rowString := ""
	    
	    for _, cell := range row.Cells {
		  
		  rowString=rowString+cell.Text+","
	    }
	    
	    outputcontent=append(outputcontent,rowString)
	}
	fmt.Printf("writing   %s", sheet.Name+".csv\n")
	if err := writeLines(outputcontent,sheet.Name+".csv" ); err != nil {
		fmt.Printf("writing: %s", err)
		var yes string
		fmt.Scan(&yes)
		return
	}
    }
	    
}