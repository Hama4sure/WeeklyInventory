package main

import (
	"fmt"
	"strings"
	"testExcel/Calculate"
	"time"
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

	oldItem := GetItemFromDcol(f) //把位於D欄的舊資料佔存在切片內
	ClearBcol(f)                  //把B欄歸0

	for i := 0; i < len(oldItem); i++ { //將切片內容導入B欄
		Replace(f, i, oldItem[i])
	}

	var str string
	fmt.Printf("請輸入，並以','做為結尾:") //使用者輸入資料
	fmt.Scan(&str)

	a := []string{}
	input := strings.Split(str, ",")
	for _, items := range input { //解析分類資料。數字=數量 文字=物品類別 ','=分隔符號
		num := ""
		item := ""
		for _, v := range items {
			if unicode.IsNumber(v) || v == '.' {
				num += string(v)
			} else {
				item += string(v)
			}
		}

		UpdateInventory(f, item, num, &a)
		num = ""
		item = ""

	}
	t := time.Now()
	dateStr := t.Format("01/02")
	UpdateInventory(f, dateStr, "", nil)
	f.SetCellValue("Sheet1", "H8", a)

	err = f.SaveAs("inventory.xlsx") //儲存變更
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("succeed")
}

func UpdateInventory(f *excelize.File, name string, num string, a *[]string) { //經解析後的單筆資料分類導入指定的欄位
	switch name {
	case "沐浴乳":
		f.SetCellValue("Sheet1", "D4", num)
		Calculate.Orderinventory(f, name, "D4", num, a)
	case "洗髮精":
		f.SetCellValue("Sheet1", "D5", num)
		Calculate.Orderinventory(f, name, "D5", num, a)
	case "酒精":
		f.SetCellValue("Sheet1", "D6", num)
		Calculate.Orderinventory(f, name, "D6", num, a)
	case "擦手紙":
		f.SetCellValue("Sheet1", "D7", num)
		Calculate.Orderinventory(f, name, "D7", num, a)
	case "洗手乳":
		f.SetCellValue("Sheet1", "D8", num)
		Calculate.Orderinventory(f, name, "D8", num, a)
	case "大捲衛生紙":
		f.SetCellValue("Sheet1", "D9", num)
		Calculate.Orderinventory(f, name, "D9", num, a)
	case "小捲衛生紙":
		f.SetCellValue("Sheet1", "D10", num)
		Calculate.Orderinventory(f, name, "D10", num, a)
	case "大垃圾袋":
		f.SetCellValue("Sheet1", "D11", num)
		Calculate.Orderinventory(f, name, "D11", num, a)
	case "小垃圾袋":
		f.SetCellValue("Sheet1", "D12", num)
		Calculate.Orderinventory(f, name, "D12", num, a)
	case "套房咖啡包":
		f.SetCellValue("Sheet1", "D13", num)
		Calculate.Orderinventory(f, name, "D13", num, a)
	case "套房餅乾":
		f.SetCellValue("Sheet1", "D14", num)
		Calculate.Orderinventory(f, name, "D14", num, a)
	case "套房牙刷":
		f.SetCellValue("Sheet1", "D15", num)
		Calculate.Orderinventory(f, name, "D15", num, a)
	case "套房棉花棒":
		f.SetCellValue("Sheet1", "D16", num)
		Calculate.Orderinventory(f, name, "D16", num, a)
	case "化妝棉":
		f.SetCellValue("Sheet1", "D17", num)
		Calculate.Orderinventory(f, name, "D17", num, a)
	case "綠茶":
		f.SetCellValue("Sheet1", "D18", num)
		Calculate.Orderinventory(f, name, "D18", num, a)
	case "髮圈":
		f.SetCellValue("Sheet1", "D19", num)
		Calculate.Orderinventory(f, name, "D19", num, a)
	default:
		f.SetCellValue("Sheet1", "D3", name)
	}
}

func ClearBcol(f *excelize.File) { //將位於B欄的舊資料移除，以0替代所有資料
	col := "B"
	value := 0
	for i := 3; i < 20; i++ {
		t := fmt.Sprintf("%s%v", col, i)
		f.SetCellValue("Sheet1", t, value)
	}

}

func Replace(f *excelize.File, index int, item string) { //將位於D欄的資料放置B欄取代
	col := "B"
	index += 3
	t := fmt.Sprintf("%s%v", col, index)
	f.SetCellValue("Sheet1", t, item)
}

func GetItemFromDcol(f *excelize.File) []string { //將D欄的資料儲存起來供Replace()使用
	oldItem := []string{}
	for i := 3; i < 20; i++ {
		col := "D"
		t := fmt.Sprintf("%s%v", col, i)
		cell, err := f.GetCellValue("Sheet1", t)
		if err != nil {
			fmt.Println(err)
			break
		}
		oldItem = append(oldItem, cell)
	}
	return oldItem
}
