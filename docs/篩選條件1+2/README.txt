實作步驟：

[篩選條件一]
"查詢銀行存款及貸方餘額" 查詢指令 產生 "deposit_credit" 資料表，供後續所有查詢使用
"同一日同一交易對手交易筆數" 查詢指令 產生 "date_counterparty_count" 資料表，"標記筆數異常分錄(同一天)" 查詢指令 產生 "day_error_count" 資料表
"同一日同一交易對手交易金額" 查詢指令 產生 "date_counterparty_sum" 資料表，"標記金額異常分錄(同一天)" 查詢指令 產生 "day_error_sum" 資料表

使用’產出7+2天迴圈’模組產出”week7+2cycle”資料表，"week7+2_count&sum" 查詢指令 產生 "week_countandsum" 資料表
"標記筆數異常分錄(週)" 查詢指令 產生 "week_error_count" 資料表，"標記金額異常分錄(週)" 查詢指令 產生 "week_error_sum" 資料表

使用’產出30+8天迴圈’模組產出”month30+8cycle”資料表，"month30+8_count&sum" 查詢指令 產生 "month_countandsum" 資料表
"標記筆數異常分錄(月)" 查詢指令 產生 "month_error_count" 資料表，"標記金額異常分錄(月)" 查詢指令 產生 "month_error_sum" 資料表

[篩選條件二]
"1-1_same_amount_for_all_account" 查詢指令 產生 "same_amount_for_all_account" 資料表
"1-1-1_same_amount_for_all_account_dollarDesc" 查詢指令 產生 "same_amount_for_all_account_dollarCalc" 資料表，供 "1-1-1_same_amount_for_all_account_dollarDesc" 畫圖使用
"1-1-2_same_amount_for_all_account_frequencyDesc" 查詢指令 產生 "same_amount_for_all_account_frequencyCalc" 資料表，供 "1-1-2_same_amount_for_all_account_frequencyDesc" 畫圖使用

"1-2_same_amount_for_selected_account" 查詢指令 產生 "same_amount_for_selected_account" 資料表
"1-2-1_same_amount_for_selected_account_dollarDesc" 查詢指令 產生 "same_amount_for_selected_account_dollarCalc" 資料表，供 "1-2-1_same_amount_for_selected_account_dollarDesc" 畫圖使用
"1-2-2_same_amount_for_selected_account_frequencyDesc" 查詢指令 產生 "same_amount_for_selected_account_frequencyCalc" 資料表，供 "1-2-2_same_amount_for_selected_account_frequencyDesc" 畫圖使用

"2-1_dollar_err" 查詢指令 產生 "dollar_err" 資料表，供 "3-1_明細帳_2_dollar_err" 查詢指令 使用
"2-2_frequency_err" 查詢指令 產生 "frequency_err" 資料表，供 "3-2_明細帳_2_frequency_err" 查詢指令 使用

"3-1_明細帳_2_dollar_err" 查詢指令 將 "dollar_err" 資料表 併回 "明細帳_2"，產生成果一 "明細帳_2_dollar_err" 資料表
"3-2_明細帳_2_frequency_err" 查詢指令 將 "frequency_err" 資料表 併回 "明細帳_2"，產生成果二 "明細帳_2_frequency_err" 資料表
