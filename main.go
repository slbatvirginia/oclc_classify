package main
// Read in an xcel spreadsheet, use a named or numbered column value 
//	to form a query to OCLC classify service to get value of
//	lcc call number.
//	Add that value to spreadsheet in OCLC-SFA column and
//	write the new spreadsheet out in new file

import (
	"flag"
	"time" 
	"fmt"
	"strings" 
	"net/http"
	"net/url"
	"io/ioutil"
	"os"
	"encoding/xml"
	"github.com/tealeg/xlsx"
)

var OCLCSite string = "http://classify.oclc.org/classify2/Classify?"

type MostRecent struct {
	Sfa string `xml:"sfa,attr"`
}
type Work struct {
	Owi string `xml:"owi,attr"`
	Wi string `xml:"wi,attr"`
}
type Response struct {
	Code string `xml:"code,attr"`
}
type Input struct {
	InType string `xml:"type,attr"`
	Value string `xml:",chardata"`
}
type OCLCResp struct {
	XMLName xml.Name `xml:"classify"`
	Response Response `xml:"response"`
	Input Input `xml:"input"`
	Works []Work `xml:"works>work"`
	Recommendations MostRecent `xml:"recommendations>lcc>mostRecent"`
}


func OCLCQuery(keyType string, keyValue string) (OCLCResp, string) {
//	accept query parameters and query OCLC classify site
//		return the response as byte array
	parsedResp := OCLCResp{}
	qs := url.Values{}
	qs.Add(keyType, keyValue)
	qs.Add("summary", "true")
	queryResp, err := http.Get(OCLCSite + qs.Encode())
	defer queryResp.Body.Close()
	if err != nil {
		fmt.Println(err)
	}
	xmlResp, err := ioutil.ReadAll(queryResp.Body)
	if err != nil {
		fmt.Println("bad http response from OCLC")
	}
	xml.Unmarshal(xmlResp,&parsedResp)
	return parsedResp, parsedResp.Response.Code
}

func OCLCRespReader(keyType string, keyValue string) OCLCResp {
//	OCLC returns spurious errors that often correct themselves on subsequent tries
//		Let's try three times if we have to.
	goodResp := false
	var resp OCLCResp
	var rcode string = "none"
	
	for i := 0; i < 5; i++ {
		resp, rcode = OCLCQuery(keyType, keyValue)
		switch rcode {
			case "0": goodResp = true
			case "4": goodResp = true
			case "2": goodResp = true
		}
		if goodResp {
			break
		}
		time.Sleep(3 * time.Millisecond)
	}
	return resp
}
	

func SfaReader(colType string, inValue string) string {
	sfa := ""
	queryResp := OCLCRespReader(colType, inValue)
	if  strings.Compare(queryResp.Response.Code,"0") == 0 {
	// single work summary - done
		sfa = queryResp.Recommendations.Sfa
	} else {
		if  strings.Compare(queryResp.Response.Code,"4") == 0 {
		// multiple works - try to use first work fields for second query
		// first try owi of first work
			owi := queryResp.Works[0].Owi
			wi := queryResp.Works[0].Wi
			queryResp = OCLCRespReader("oclc",owi)
			if  strings.Compare(queryResp.Response.Code,"0") == 0 {
			// got a hit - done
				sfa = queryResp.Recommendations.Sfa
			} else {
			// failed - try the wi of first work
				queryResp = OCLCRespReader("oclc",wi)
				if  strings.Compare(queryResp.Response.Code,"0") == 0 {
				// finally - done
					sfa = queryResp.Recommendations.Sfa
				}
			}
		} 
	}
	
	return sfa
}


func matchColHeader(headerRow xlsx.Row, matchName string) int {
	var colNbr int = -1
	for nCell, iCell := range headerRow.Cells {
		iValue, err := iCell.String()
		if err != nil {
			iValue = "none"
		}

		if strings.Compare(iValue,matchName) == 0 {
			fmt.Printf("%s translated to col %d\n",matchName,nCell)
			colNbr  = nCell
		}
	}
	return colNbr
}
	

func main() {
//      input params processing
	var version string = "20170321"
	current_time:= time.Now().Local()
	version = current_time.Format("20060102");
	fmt.Printf("compiled %s\n",version);
	var inFile, outFile, colName, colType string
	flag.StringVar(&inFile, "infile", "./in.xlsx", "input file")
	flag.StringVar(&outFile, "outfile", "./out.xlsx", "output file")
	flag.StringVar(&colName, "colname", "none", "xcel col hdr to match (overides pos)")
	flag.StringVar(&colType, "coltype", "issn", "name of OCLC query parameter")
	var colNbr int = -1

	colPosPtr :=flag.Int("colpos", 2, "xcel column position to use ")
	flag.Parse()
	xfilein, err := xlsx.OpenFile(inFile)
	if err != nil {
		fmt.Printf("open fail %s for %s\n", err, inFile)
		os.Exit(2)
	}
	xfileout := xlsx.NewFile();
	osheet, err := xfileout.AddSheet("OCLCClassifyOut")
	if err != nil {
		fmt.Printf("xcel fail %s",err)
		os.Exit(2)
	}

//  	if passed in column header, override position with matching column
	if strings.Compare(colName,"none") != 0 {
		fmt.Printf("matching on col heading %s\n",colName)
		headerRow := xfilein.Sheets[0].Rows[0]
		colNbr = matchColHeader(*headerRow,colName)	
	}
	
//	if no match on header, use the number position 
	if colNbr < 0 {
		colNbr = *colPosPtr;
	}
		
	fmt.Printf("using column %d\n",colNbr)


//      Process rows in spreadsheet, create new spreadsheet
	got := 0
	tried := 0
	for  nRow, iRow := range xfilein.Sheets[0].Rows {
		oRow := osheet.AddRow()
		var iStyle *xlsx.Style
		var sfaValue string = ""
		for nCell, iCell := range iRow.Cells {
			oCell := oRow.AddCell()
//			get style and value from input and copy it to output
			iValue, _ := iCell.String()
			iStyle = iCell.GetStyle()
			oCell.SetStyle(iStyle)
			oCell.SetString(iValue)

		    	if nCell == colNbr {
				if strings.Compare(iValue,"") != 0 {
					tried = tried + 1
					sfaValue = SfaReader(colType,iValue)
					if strings.Compare(sfaValue,"") != 0 {
						got = got + 1
					}	
				}
			}
		}
			
	
//  		Add in whatever sfa value you have to output xcel sheet
		sfaValueCell := oRow.AddCell()
		sfaValueCell.SetStyle(iStyle)
		if nRow == 0 {
			sfaValueCell.SetString("OCLC-SFA")
		} else {	
			sfaValueCell.SetString(sfaValue)
		}
	}
// 	Save the new spreadsheet containing sfa values
	fmt.Printf("tried %d, got values for %d\n",tried,got)
	err =  xfileout.Save(outFile)
	if err != nil {
		fmt.Printf("Error %s saving file %s\n",err,outFile)
	}
	
}

//!-
