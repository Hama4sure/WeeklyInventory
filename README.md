# WeeklyInventory  
Brief Description: By combining Excel and Golang, this program records inventory (items and quantities) and views the previous inventory record.

This is a program written in the Go language, whose main function is to read the old item data from an Excel file, import the new item data into the Excel file, and remove the old data in column B.

The program will ask the user to input item data in the format of "item category, quantity", where the quantity must be a number. The program will parse the user's input data and import the quantity into the corresponding column of the item category in the Excel file. For example, if the user inputs "3shampoo", the program will import the quantity 3 into the D5 column (shampoo column) of the Excel file.

The main logic of the program is implemented in the PutInto() function, which imports the quantity into the corresponding column of the Excel file based on the item category. The program also includes the Remove() function, which removes the old data in column B, and the Replace() function, which places the old data in column D into column B to replace the original data.

Finally, the program will save the changes to the Excel file and display a success message.

User Guide: Example of user input data: 2023/02/03,3alcohol,7hand soap,12tissue paper,4body wash,8shampoo,2cotton pad,6big toilet paper,9small toilet paper. The date format is YYYY/MM/DD. Each item starts with a number followed by the item name, separated by a comma.

Feedback and Suggestions: If there are any issues or suggestions regarding the program, please feel free to contact me at tioka.speed@gmail.com.

# 每周庫存紀錄
簡易描述: 藉由串接Excel以及Golang，來完成紀錄庫存(物品以及數量)以及檢視上回庫存紀錄

這是一個使用 Go 語言編寫的程式，主要的功能是從 Excel 檔案中讀取舊的物品資料，然後將新的物品資料導入 Excel 檔案中，並將 B 欄位的舊資料移除。

程式會要求使用者輸入物品資料，資料格式為：「物品類別,數量」，其中數量必須是數字。程式會解析使用者輸入的資料，並根據物品類別將數量導入 Excel 檔案中的相應欄位。例如，如果使用者輸入「洗髮精,3」，程式會將數量 3 導入 Excel 檔案中的 D5 欄位（洗髮精欄位）。

程式的主要邏輯是在 PutInto() 函數中實現的，該函數根據物品類別將數量導入 Excel 檔案中的相應欄位。程式還包括 Remove() 函數，用於將位於 B 欄的舊資料移除，以及 Replace() 函數，用於將位於 D 欄的舊資料放置 B 欄中以替換掉原有的資料。

最後，程式會將變更儲存到 Excel 檔案中並顯示成功訊息。

使用指南: 使用者輸入的資料(範例): 2023/02/03,3酒精,7洗手乳,12擦手紙,4沐浴乳,8洗髮乳,2化妝棉,6大捲衛生紙,9小捲衛生紙,
日期格式為西元/月月/日日。每一個物品的開頭數字為數字並直接接上庫存物品名，並以","連結文字以及結尾

建議以及回饋: 如有程式上的問題或是建議的話，以下是我的Email，歡迎與我聯繫
tioka.speed@gmail.com
