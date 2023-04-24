package Calculate

import (
	"fmt"
	"strconv"
	"testExcel/types"

	"github.com/xuri/excelize/v2"
)

func Orderinventory(f *excelize.File, item, col, amount string, a *[]string) {
	amountfloat, _ := strconv.ParseFloat(amount, 64)
	if amountfloat <= types.Centralmap[item] {
		sentance := fmt.Sprintf("%s一箱\n", item)
		*a = append(*a, sentance)
	}
}
