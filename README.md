# Workshop guide page: setup

The page (`index.html`) is static and goes on GitHub Pages. Shared entries live in a Google Sheet behind a small Apps Script (`apps-script/Code.gs`). Setup takes about 20 minutes.

## 1. The Google Sheet and script (10 min)

1. Create a new Google Sheet, for example "Workshop board".
2. In the sheet, open **Extensions → Apps Script**.
3. Delete the sample code, and paste in all of `apps-script/Code.gs`.
4. Click the gear (**Project Settings**), scroll to **Script properties**, and add a property `FACILITATOR_KEY` whose value is a word only you and the facilitators know. Keeping the key there means `Code.gs` can sit in a public repository without revealing it.
5. Go back to the editor (**<>**), pick `setup` from the function menu at the top and click **Run**. Google asks for permission the first time. Approve it. The script then creates the `entries` tab with headers.
6. Click **Deploy → New deployment**. Choose the type **Web app**. Set *Execute as* to **Me** and *Who has access* to **Anyone**. Click **Deploy**.
7. Copy the web app URL. It ends in `/exec`.

If you later change the script, use **Deploy → Manage deployments → Edit → Version: New version**. That keeps the same URL.

## 2. The page (5 min)

1. Open `index.html` and find `CONFIG` near the bottom.
2. Paste the URL into `apiUrl`.
3. Adjust `groups` (two or three) and the `agenda` times if needed.
4. Push to a GitHub repository and switch on **Settings → Pages → Deploy from branch → main / root**.

You can also test without editing the file. Append `?api=YOUR_EXEC_URL` to the page address. With no URL at all, the page runs in demo mode and saves only in your own browser.

## 3. Test before the day (5 min)

1. Open the page on two devices.
2. Add an entry on one, and check that it appears on the other within 20 seconds (or press Refresh).
3. Click **Me too** and check that the count rises.
4. Click **Facilitator mode** at the bottom, enter the key, and mark the entry as picked.
5. Open the sheet and delete the test rows.

## What to know

- **Anyone with the page link can write to the sheet.** That is fine for one day. Afterwards, archive the deployment (Deploy → Manage deployments → Archive).
- **Votes are only guarded per browser.** Someone could vote twice from two browsers. For a workshop of 16 colleagues, that risk is acceptable.
- **The facilitator key is a lock against accidents, not a security measure.**
- **The sheet sits in your Google account.** If the faculty prefers Microsoft 365, the same pattern is possible with a Microsoft Form and an embedded Excel, but participants would lose the voting.
