# Packing list preferences

- Always use A4 paper for packing lists.
- Use one clearly separated bordered table per order. Show the Order ID once as the table header and keep all products for that order underneath it.
- Keep each entire order table together on one page.
- Include Employee:__________________ after the Order ID/Order Number, followed by the blank Number and Time underlines inside each order table header. Staff write their own number to link the packing list to the printed waybill, and the packing completion time.
- Include product name, variation, and number of pieces. Preserve source batch grouping.
- At the end of every generated packing document, add a product picking summary after a page break. Apply the saved Define Name alternatives independently to product name and variation; blank alternatives retain their original fields. Combine matching final product names and variations only within that generated document. Preserve original source data and fingerprints in processing history, plus a separate immutable snapshot of the printed names.
- Use light orange order headers and light peach column headers with dark text for SPX/FLASH. All product line-item rows use light gray shading.
- For Lazada, use the Order Number as the order header and omit Tracking ID. Use blue order and column headers with dark text and light gray product line-item rows, while keeping the same A4 order grouping and handwritten Number and Time fields.
- Combine Lazada and TikTok in one document for each processing batch, using the blue theme and clear platform section labels. Retain Lazada Order Number and TikTok Order ID as appropriate. Use a shared `LAZADA_TIKTOK` filename sequence. Mark earlier standalone outputs as superseded in the processing log when a combined replacement includes their orders.
- Save all finished packing documents together in a `000 Output` subfolder of the existing dated source folder, for example `20260926/000 Output`. Do not create separate carrier output folders or a parallel year/month/day directory tree.
- End each packing document filename with a sequence suffix before `.pdf`. For each output group and date (SPX, FLASH, or combined LAZADA_TIKTOK), use one above the highest numbered PDF or Word file still present in `000 Output`, or `_1` if none remain. Reuse a deleted trailing number but do not fill a lower gap: existing `_1, _2` produces `_3`; existing `_1, _3` produces `_4`. New exports are PDF only; existing Word files still count toward the sequence. Never overwrite an existing file.
- Maintain `Processed_Files.txt` in that dated `000 Output` folder, recording each completed output and its source PDFs. Check this log before processing later arrivals, and include only newly added unprocessed files unless the user requests a reprint or revision. Review changed source files explicitly rather than silently treating them as new or unchanged.
- The existing Windows program is authorized. Save completed version checkpoints locally in Git, with a verified bundle in Git/vVERSION. Preserve earlier version folders; do not upload without authorization.

- Define Name stores uppercase single-line fields, trims trailing spaces, and matches original pairs without case sensitivity. Confirm before merging rows with equal nonblank alternative product AND variation. Keep the first selected base and every linked original pair; later PDF scans must recognize all links. Excel round trips must preserve and permit editing those linked originals.

- Shopee continuation labels with the same Order ID and consecutive printed item numbers belong to one order. Reject missing, overlapping or duplicate sections instead of silently dropping products.
- Each Shopee PDF has Auto/SPX/FLASH courier selection. Auto takes the nearest exact SPX or FLASH folder first, then filename/path hints. Explicit per-file selection overrides Auto and persists until reset or PDF contents change. Lazada/TikTok cells show a disabled dash.
- Scan compatible PDFs anywhere under the selected dated folder except 000 Output. Hide unrecognized/no-readable-text PDFs; keep genuine errors in recognized packing lists visible for review.

- Main screen labels have no step numbering. Generate exports PDF only with no format-selection dialog; the public engine must reject Word export requests. Existing PDF and Word outputs retain their history and numbering coverage.
- Before generating repeated orders (earlier dated history or overlapping source PDFs), show the requested warning with a large centered No button as Enter/Escape default. Below it, place smaller light-gray borderless text reading Continue to Generate the Repeated Orders, centered within the No button width. Matching orders print once per run. Backend confirmation is tied to the selection and output format.
- Define Name opens locked each time with icon-only closed/open padlocks. Locked mode permits search/view but disables changes and Excel editing. Unlock explicitly; save before relocking.
- Use [     ] for Done boxes in both packing tables and picking summaries. Show dynamic totals for checked PDFs, distinct order IDs and pieces below the main list, with the selected platforms only.

- Completed source rows show only Process again in Status. Retain missing-output and review errors when applicable. Use light gray product rows in both order tables and picking summaries; preserve platform-colored headers and summary totals.
