package main

import (
	"fmt"
	"strconv"
	"unicode"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("inventory.xlsx") //開啟excel檔案
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()

	last := SaveOldDatas(f) //把位於D欄的舊資料佔存在切片內
	Remove(f)               //把B欄歸0

	for i := 0; i < len(last); i++ { //將切片內容導入B欄
		Replace(f, i, last[i])
	}

	var str string
	fmt.Printf("請輸入，並以','做為結尾:") //使用者輸入資料
	fmt.Scan(&str)
	num := ""
	item := ""

	for k, v := range str { //解析分類資料。數字=數量 文字=物品類別 ','=分隔符號
		if k < 10 { //時間格式為2023/03/17, 十位數
			item += string(v)
			continue
		}

		if unicode.IsNumber(v) {
			num += string(v)

		} else if v == ',' {
			c, _ := strconv.Atoi(num)
			PutInto(f, item, c)
			num = ""
			item = ""

		} else {
			item += string(v)
		}

	}

	err = f.SaveAs("inventory.xlsx") //儲存變更
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("succeed")
}

func PutInto(f *excelize.File, name string, num int) { //經解析後的單筆資料分類導入指定的欄位
	switch name {
	case "沐浴乳":
		f.SetCellValue("Sheet1", "D4", num)
	case "洗髮精":
		f.SetCellValue("Sheet1", "D5", num)
	case "酒精":
		f.SetCellValue("Sheet1", "D6", num)
	case "擦手紙":
		f.SetCellValue("Sheet1", "D7", num)
	case "洗手乳":
		f.SetCellValue("Sheet1", "D8", num)
	case "大捲衛生紙":
		f.SetCellValue("Sheet1", "D9", num)
	case "小捲衛生紙":
		f.SetCellValue("Sheet1", "D10", num)
	case "大垃圾袋":
		f.SetCellValue("Sheet1", "D11", num)
	case "小垃圾袋":
		f.SetCellValue("Sheet1", "D12", num)
	case "套房咖啡包":
		f.SetCellValue("Sheet1", "D13", num)
	case "套房餅乾":
		f.SetCellValue("Sheet1", "D14", num)
	case "套房牙刷":
		f.SetCellValue("Sheet1", "D15", num)
	case "套房棉花棒":
		f.SetCellValue("Sheet1", "D16", num)
	case "化妝棉":
		f.SetCellValue("Sheet1", "D17", num)
	case "綠茶":
		f.SetCellValue("Sheet1", "D18", num)
	case "髮圈":
		f.SetCellValue("Sheet1", "D19", num)
	default:
		f.SetCellValue("Sheet1", "D3", name)
	}
}

func Remove(f *excelize.File) { //將位於B欄的舊資料移除，以0替代所有資料
	col := "B"
	value := 0
	for i := 3; i < 20; i++ {
		t := fmt.Sprintf("%s%v", col, i)
		f.SetCellValue("Sheet1", t, value)
	}

}

func Replace(f *excelize.File, i int, item string) { //將位於D欄的資料放置B欄取代
	col := "B"
	i += 3
	t := fmt.Sprintf("%s%v", col, i)
	f.SetCellValue("Sheet1", t, item)
}

func SaveOldDatas(f *excelize.File) []string { //將D欄的資料儲存起來供Replace()使用
	last := []string{}
	for i := 3; i < 20; i++ {
		col := "D"
		t := fmt.Sprintf("%s%v", col, i)
		cell, err := f.GetCellValue("Sheet1", t)
		if err != nil {
			fmt.Println(err)
			break
		}
		last = append(last, cell)
	}
	return last
}
