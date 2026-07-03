# 新埔八街28號3樓套房出租管理系統

靜態版出租管理系統原型，支援 GitHub Pages 部署。資料策略是：

- 文字與數字資料回寫 Google Sheet 後臺主表。
- 圖片、PDF、租約、收據等附件分類上傳到 Google Drive 子資料夾。
- 前端使用 Google Identity Services 取得 Drive 與 Sheets 權限。

## Google 登入設定

正式登入需要 Google Cloud OAuth Client ID。建立 OAuth 用戶端後，把 Client ID 填入 `config.js`：

```js
window.RENTAL_GOOGLE_CLIENT_ID = "你的 OAuth Client ID";
```

OAuth 授權來源需加入 GitHub Pages 網址，例如：

```text
https://craigyu-phd.github.io
```

授權重新導向 URI 不是必要項，因為本系統使用瀏覽器端 token flow。

## Drive 工作區

- 主資料夾：`新埔八街28號3樓-套房出租管理系統`
- 後臺主表：`新埔八街28號3樓-後臺主表`
- 附件子資料夾：租約、臺電帳單、電表照片、維修附件、收款憑證、租客附件
