# Quick Text – Word add-in

A task pane in Word with your own list of text snippets. Put the cursor where you want text, click a snippet, and it is inserted there (it replaces any text you have selected).

Features: add / edit / delete snippets, groups, search, backup & restore (copy/paste JSON).

## Run it (Windows, Word desktop)

You need [Node.js](https://nodejs.org) installed once.

1. Open a terminal in this folder.
2. One time only – install a trusted local certificate (accept the prompt):
   ```
   npx office-addin-dev-certs install
   ```
3. Start the local server and leave the window open while you use Word:
   ```
   node serve.js
   ```
4. Load the add-in into Word (shared-folder method):
   1. Create a folder, e.g. `C:\WordAddins`, copy `manifest.xml` into it, right-click the folder → Properties → Sharing → Share. Copy the network path (`\\YOUR-PC\WordAddins`).
   2. In Word: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs. Paste the path into "Catalog Url", click **Add catalog**, tick **Show in Menu**, OK. Restart Word.
   3. In Word: Insert → My Add-ins → **Shared Folder** tab → select **Quick Text** → Add.
5. Home tab → **Quick Text** button opens the pane.

### Word on the web (quick test)
Open a document in Word for the web → Insert → Add-ins → Advanced / "Upload My Add-in" → choose `manifest.xml`. (The local server must be running and the certificate trusted in your browser.)

### Mac
Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`, run the same two commands above, restart Word, then Insert → My Add-ins → Developer Add-ins.

## Using it without keeping a server running
The add-in is just static files. You can host this folder on any free HTTPS static host (GitHub Pages, Netlify, Cloudflare Pages), then replace every `https://localhost:3000` in `manifest.xml` with your site address and load the manifest as above.

## Notes
- Snippets are saved in the add-in's own browser storage on that computer. Use **Backup** to copy them out, and **Import** to restore or move them to another computer.
- Line breaks in a snippet become new paragraphs when inserted.
- Plain text only for now. Formatting (bold, tables) would need the snippet stored as HTML/OOXML – can be added.

## Features (v2)
- **Folder tabs**: All, ★ Favorites, and your own 📁 folders (+ Folder to add; Rename / Delete when a folder tab is open).
- **Favorites**: tap ☆ on a snippet; see them all in the Favorites tab.
- **Pin**: tap 📌 to keep a snippet at the top of its list.
- **Multiple tags**: add as many tags as you like (Enter or comma); click a tag to filter.
- **Resizable text box**: drag the bottom-right corner of the editor box; the size is remembered.
- **Reorder**: drag the ⠿ handle up or down, or use the ↑ / ↓ buttons (arrow keys also work on the handle). Order is saved. Pinned snippets stay above unpinned ones.
- Old snippets (and their old "group") are converted automatically.
