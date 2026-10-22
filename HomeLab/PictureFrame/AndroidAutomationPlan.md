# Android share-to-display workflow

The display page checks `current.txt` every three seconds. Put the selected media file in the same folder as `index.html` and make `current.txt` contain its filename on one line, for example:

```text
current.png
```

The page supports common image formats (`jpg`, `jpeg`, `png`, `gif`, `webp`, `avif`, and `bmp`) and browser-playable video (`mp4`, `webm`, and `ogv`). Videos autoplay muted, loop, and have controls. Media is centered and contained in the browser window, so the whole image is visible without cropping. For images, use **Zoom to actual size** to inspect and scroll around a large image, then **Fit to screen** to return to the full image.

## MacroDroid outline

1. Create a macro with the **File Shared to MacroDroid** trigger. This makes MacroDroid selectable in Android's Share menu.
2. Allow the trigger to save the incoming file to a staging folder, or otherwise copy it from the trigger's shared-file path. Grant MacroDroid access to the source and destination folders when prompted.
3. Read the existing filename from `Documents/livedisplay/current.txt` (if present) and keep it in a local string variable. This identifies the previously displayed file for cleanup.
4. Copy the shared file into `Documents/livedisplay`, using its shared filename so its extension stays intact.
5. Overwrite `Documents/livedisplay/current.txt` with only the new filename. Do this after the copy so the page does not switch to a file that is still being written.
6. Delete the previous media file named in the variable if it differs from the new filename. Do not delete `index.html` or `current.txt`.

If MacroDroid cannot reliably read and write these files on the Android version in use, use an automation app with file access or a small upload endpoint on the server instead. The server must serve `current.txt` and the media files from the same directory as `index.html`.

## Existing setup

If `current.txt` is missing or unavailable, `index.html` falls back to `current.jpg` for compatibility with the existing single-image setup. To select other media types, make sure the server can serve `current.txt` as plain text.