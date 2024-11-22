# SecureChangeRequestFormatter
## 這是什麼？
這是供客戶使用的轉換腳本，能將客戶的需求表單轉換為 Secure Change 需要的格式<br>
為因應更多樣的變化，提供 `config.json` 作為關鍵字設定以提高使用彈性

## 參考資料
- [Tufin 原廠文件](https://forum.tufin.com/support/kc/latest/Content/Suite/change-request_advanced-options.htm)

## 使用方法
1. 將 `scformatter.exe` 放在資料夾中，確保 `config.json` 存在
2. 將欲轉換萃取的 Excel 檔案(`*.xlsx`) 放在相同資料夾。可以放置多個
3. 執行 `scformatter.exe`
4. 生成 `YYYY-MM-DDThh_mm_ss.xlsx` 與 `output.log` 檔案。前者為生成後的檔案；後者為紀錄檔，為終端輸出的備份。終端會顯示相關資訊，如因格是不合而「略過的項目」等。回車可關閉終端
5. 使用。複製 `output.log` 即可貼入 Tufin Secure Change Request 中。請參考[Tufin 原廠文件](https://forum.tufin.com/support/kc/latest/Content/Suite/change-request_advanced-options.htm)

## 原理
- TOS Secure Change Request 需要「來原」、「目標」、「服務」三大欄位
- 搜尋所有同資料夾下 `*.xlsx` 的所有表單，並嘗試在表單首兩行找到關鍵字作為欄位，失敗則跳過
- 關鍵字設定在 `config.json` 中
- 每列需要欄位全空會略過，必要欄位有缺漏會略過並紀錄
- 結果最終輸出到 `output.log`

## 注意
- 不支援舊版 Excel 檔案 (`*.xls`)
- 每個欄位可以填入複數比資料，可使用分號(';')、換行('\n')進行分隔，不能使用空白(' ')分隔。程式會轉化為正確格式(分號分隔)
- 服務欄位需要填入協議類型，例如 "TCP 22" 或 "UDP 514"。沒有協議只有埠口會默認設定為 "TCP"
- 因客需表單設定，「行動」被拆分為「刪除」與「新增」兩個欄位。可以皆空(原廠表格允許)但不濃同時有值。「行動」的類別只支援「刪除」與「新增」，也就是純白名單，只有`accept`, `remove`，沒有 `drop`
- 「來原」、「目標」、「服務」為必填欄位，
- 因會檢查欄位會檢查 Row1, Row2，項目則由 Row 2 開始檢查，因此 Row 2 可能產生誤報，還請忽略

## 輸出範例
![](/image/example.png)