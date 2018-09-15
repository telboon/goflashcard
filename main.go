package main

import (
   "fmt"
   "strconv"
   "math"
   "os"
   "bufio"
   "math/rand"
   "github.com/360EntSecGroup-Skylar/excelize"
   "time"
)

func convertCell(columnNo int, rowNo int) string {
   var highestPow int
   var tempCol int
   tempCol = 1

   colStr := ""
   rowStr := strconv.Itoa(rowNo)
   colIndex := " ABCDEFGHIJKLMNOPQRSTUVWXYZ"

   for ;columnNo/tempCol >0; highestPow+=1 {
      tempCol = int(math.Pow(26, float64(highestPow)))
   }
   highestPow-=2

   tempCol = columnNo
   for i:=highestPow; i>=0; i-- {
      indexCol:= tempCol / int(math.Pow(26,float64(i)))
      tempCol -= int(math.Pow(26,float64(i))) * indexCol
      colStr = colStr + colIndex[indexCol:indexCol+1]
   }

   return colStr+rowStr
}

func main() {
   highestRow := 0
   dataFile, err := excelize.OpenFile("./data.xlsx")
   inReader := bufio.NewReader(os.Stdin)

   if err!= nil {
      fmt.Println(err)
      return
   }
   for highestRow=1;dataFile.GetCellValue("flashdata", convertCell(1,highestRow))!="";highestRow+=1 {
   }

   highestRow-=1

   tempInput:= ""
   fmt.Println("Type <enter> to continue, 'q' to quit.")
   tempInput, _ = inReader.ReadString('\n')
   rand.Seed(time.Now().UTC().UnixNano())
   for ;tempInput!="q\n"; {
      row:= rand.Int() % highestRow
      row+=1

      fmt.Println("Question: "+dataFile.GetCellValue("flashdata", convertCell(1,row)))
      tempInput, _ = inReader.ReadString('\n')
      if tempInput=="q\n" {
         break
      }
      fmt.Println("Answer: "+dataFile.GetCellValue("flashdata", convertCell(2,row)))
      fmt.Println()
   }

}
