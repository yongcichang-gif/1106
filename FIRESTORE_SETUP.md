# Firestore 設定

Daymark 使用 Google 登入的 Firebase UID 分開保存任務。新帳號第一次登入時，會把目前的示範任務複製到該帳號；之後新增、完成或刪除任務會同步到 Firestore。

## 啟用 Firestore

1. 開啟 Firebase 專案 `daymark-f9a4f` 的 Firebase Console，確認 Authentication 的 Google 登入供應商已啟用，並在「Authorized domains」加入 `yongcichang-gif.github.io`。
2. 前往「Build」→「Firestore Database」，建立資料庫。
3. 在專案資料夾執行以下命令發布本專案的安全規則：

   ```powershell
   firebase login
   firebase deploy --only firestore:rules
   ```

   `.firebaserc` 已將預設專案設為 `daymark-f9a4f`。也可以把 `firestore.rules` 的內容貼到 Firebase Console 的 Firestore「規則」頁籤，然後發布。

4. 將更新後的 `index.html`、`firebase.json`、`.firebaserc` 和 `firestore.rules` 發布到 GitHub Pages，網站才會開始使用雲端任務。

規則只允許已登入使用者讀寫 `users/{自己的 UID}/tasks`。登出後，介面會回到示範資料；登入後載入該帳號的雲端資料。
