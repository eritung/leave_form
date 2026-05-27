# 艾迪英特請假單 Streamlit 產生器

## 功能
- 使用原始 `艾迪英特請假單.xlsx` 作為模板。
- 使用者在 Streamlit 頁面填寫請假人、代理人、假別、事由、起訖時間與請假天數。
- 送出後會同步填入「公司留存」與「請假者留存」兩聯。
- 可直接下載保留原版型的 xlsx。

## 本機執行
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 上架
1. 將以下三個檔案放在同一個 GitHub repo：
   - `app.py`
   - `requirements.txt`
   - `艾迪英特請假單.xlsx`
2. 到 Streamlit Cloud 建立 App。
3. Main file path 填入：`app.py`
4. Deploy 後即可使用。

## 欄位對應
- 申請日期：上半部 I4/K4/M4，下半部 I25/K25/M25
- 請假人：上半部 D5，下半部 D26
- 代理人：上半部 K5，下半部 K26
- 假別勾選：上半部第 8 列，下半部第 29 列
- 事由：上半部 D9，下半部 D30
- 開始時間：上半部 C11/E11/G11/I11/K11，下半部 C32/E32/G32/I32/K32
- 結束時間：上半部 C12/E12/G12/I12/K12，下半部 C33/E33/G33/I33/K33
- 請假天數：上半部 M11，下半部 M32
