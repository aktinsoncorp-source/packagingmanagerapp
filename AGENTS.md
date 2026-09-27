# Packing list preferences

- Always use A4 paper for packing lists.
- Use one clearly separated bordered table per order. Show the Order ID once as the table header and keep all products for that order underneath it.
- Keep each entire order table together on one page.
- Include blank Number and Time underlines inside each order table header. Staff write their own number to link the packing list to the printed waybill, and the packing completion time.
- Include product name, variation, and number of pieces. Preserve source batch grouping.
- At the end of every generated packing document, add a product picking summary after a page break. Use the original product names and variation text from the source PDFs, without shortened names or aliases. Sum pieces only for matching original product names and variations; keep different listing names, scents, sizes, grades, and product types separate. Preserve sizes within their original source text, and use a dash when the source has no variation. Each summary and its total must cover ONLY that generated document, never all output files together.
- Use light orange order headers with dark text. Use subtle pale orange or cream shading where it improves readability.
- For Lazada, use the Order Number as the order header and omit Tracking ID. Use a light blue table theme with dark text and subtle pale blue shading, while keeping the same A4 order grouping and handwritten Number and Time fields.
- Combine Lazada and TikTok in one document for each processing batch, using the blue theme and clear platform section labels. Retain Lazada Order Number and TikTok Order ID as appropriate. Use a shared `LAZADA_TIKTOK` filename sequence. Mark earlier standalone outputs as superseded in the processing log when a combined replacement includes their orders.
- Save all finished packing documents together in a `000 Output` subfolder of the existing dated source folder, for example `20260926/000 Output`. Do not create separate carrier output folders or a parallel year/month/day directory tree.
- End each packing document filename with a sequence suffix before `.docx`. For each output group and date (SPX, FLASH, or combined LAZADA_TIKTOK), use one above the highest numbered Word file still present in `000 Output`, or `_1` if none remain. Reuse a deleted trailing number but do not fill a lower gap: existing `_1, _2` produces `_3`; existing `_1, _3` produces `_4`. Never overwrite an existing Word file.
- Maintain `Processed_Files.txt` in that dated `000 Output` folder, recording each completed output and its source PDFs. Check this log before processing later arrivals, and include only newly added unprocessed files unless the user requests a reprint or revision. Review changed source files explicitly rather than silently treating them as new or unchanged.
- Do not build a reusable program until the user gives the signal.
