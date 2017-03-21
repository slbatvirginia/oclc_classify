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

func getOCLCResp(keyType string, keyValue string) []byte {
//	accept query parameters and query OCLC classify site
//		return the response as byte array
	qs := url.Values{}
	qs.Add(keyType, keyValue)
	qs.Add("summary", "true")
	resp, err := http.Get(OCLCSite + qs.Encode())
	defer resp.Body.Close()
	if err != nil {
		fmt.Println(err)
	}
	b, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		fmt.Println("bad response from OCLC")
	}
	return b
	
}

func getSfa(inValue string) string {
//
//	unmarshall the xml response 
//	parse the resulting struct to extract a call number
//		for single work responses, use recommendations>lcc>mostRecent>sfa
//		for multiple work responses, requery using owi or wi from first work
//			
	type MostRecent struct {
		Sfa string `xml:"sfa,attr"`
	}
	type Work struct {
		Owi string `xml:"owi,attr"`
		Wi string `xml:"wi,attr"`
	}
	type ResponseCode struct {
		Code string `xml:"code,attr"`
	}
	type Input struct {
		InType string `xml:"type,attr"`
		Value string `xml:",chardata"`
	}
	type OCLCresp struct {
		XMLName xml.Name `xml:"classify"`
		ResponseCode ResponseCode `xml:"response"`
		Input Input `xml:"input"`
		Works []Work `xml:"works>work"`
		Recommendations MostRecent `xml:"recommendations>lcc>mostRecent"`
	}
	XMLResp := OCLCresp{}
	respBody := getOCLCResp("issn",inValue)
	xml.Unmarshal(respBody, &XMLResp)
	TagValue := ""
	Tag := ""
	// see if you can get an sfa value from response(s)
	if  strings.Compare(XMLResp.ResponseCode.Code,"0") == 0 {
	// single work summary
		Tag = "sfa"
		TagValue = XMLResp.Recommendations.Sfa
	} else {
		if  strings.Compare(XMLResp.ResponseCode.Code,"4") == 0 {
		// multiple works 
		// first try owi of first work
			owi := XMLResp.Works[0].Owi
			wi := XMLResp.Works[0].Wi
			XMLResp := OCLCresp{}
			respBody = getOCLCResp("oclc",owi)
			xml.Unmarshal(respBody,&XMLResp)
			if  strings.Compare(XMLResp.ResponseCode.Code,"0") == 0 {
			// got a hit
				Tag = "sfa"
				TagValue = XMLResp.Recommendations.Sfa
			} else {
			// failed - try the wi of first work
				XMLResp := OCLCresp{}
				respBody = getOCLCResp("oclc",wi)
				xml.Unmarshal(respBody,&XMLResp)
				if  strings.Compare(XMLResp.ResponseCode.Code,"0") == 0 {
				// finally
					Tag = "sfa"
					TagValue = XMLResp.Recommendations.Sfa
				}
			}
		} 
	}
	
	if strings.Compare(Tag,"sfa") == 0 {
		return TagValue
	} else {
		return ""
	}
}


func matchColHeader(headerRow xlsx.Row, matchName string) int {
	var colNbr int = -1
	for nCell, iCell := range headerRow.Cells {
		iValue, err := iCell.String()
		if err != nil {
			fmt.Printf("bad header cell %d\n",nCell)
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
	var inFile, outFile, colName string
	flag.StringVar(&inFile, "infile", "./in.xlsx", "input file")
	flag.StringVar(&outFile, "outfile", "./out.xlsx", "output file")
	flag.StringVar(&colName, "colname", "none", "xcel col hdr to match (overides pos)")
	var colNbr int

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
		fmt.Printf("Using col heading %s\n",colName)
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
				sfaValue = getSfa(iValue)
				if strings.Compare(sfaValue,"") != 0 {
					got = got + 1
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
	err =  xfileout.Save(outFile)
	if err != nil {
		fmt.Printf("Error %s saving file %s\n",err,outFile)
	}
	
}

//!-
