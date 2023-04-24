package types

var Centralmap = make(map[string]float64)

func defineInventory() {
	Centralmap["沐浴乳"] = 2
	Centralmap["洗髮精"] = 2
	Centralmap["酒精"] = 2

	Centralmap["擦手紙"] = 8
	Centralmap["洗手乳"] = 2
	Centralmap["大捲衛生紙"] = 5
	Centralmap["小捲衛生紙"] = 3
	Centralmap["大垃圾袋"] = 7
	Centralmap["小垃圾袋"] = 10

	Centralmap["套房咖啡包"] = 0.5
	Centralmap["套房餅乾"] = 0.5
	Centralmap["套房牙刷"] = 0.5
	Centralmap["套房棉花棒"] = 0.5
	Centralmap["化妝棉"] = 2
	Centralmap["綠茶"] = 1

}
