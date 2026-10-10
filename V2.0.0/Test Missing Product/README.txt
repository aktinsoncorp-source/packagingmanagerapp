MISSING PRODUCT TEST - Packing List Manager V2.0.0

Close any open V2 Packing List Manager window, then double-click
Start Missing Product Test.cmd in this folder.
It opens the Python app with independent test names and settings in AppData.
The normal application's ProductNames.json and Settings.json are not used.

The test date is 10 October 2026 (20261010).
If choosing the folder manually, choose this Test Missing Product folder,
set the packing date to 10 October 2026, and click Refresh.
Use the test launcher first to keep your normal definitions separate.

There is exactly one synthetic FLASH order:
Order ID: TESTNOSKU20261010A
Ecommerce Name: TEST PRODUCT - MOONLIGHT VANILLA
Ecommerce Variation: 50 mL
Quantity: 1
Seller SKU: blank
Prices: unknown (blank)

The product intentionally does not exist in the current product catalog.
It is for testing only. Do not ship it or use it as a real customer order.

Try Later/ignoring the names prompt, then Generate Packing Lists. The test
output should retain the ecommerce name and variation. The reminder should
continue to direct attention to the missing SKU and names.

Then open the highlighted product, explicitly unlock editing and its row,
and enter a Product Name and Size. Save names, then generate again if desired.
A missing-SKU reminder may remain until its ecommerce SKU has been corrected.
Test output is saved beneath 20261010/000 Output.

Keep the existing source PDF; use completed-document/reprint controls to
repeat a generated order. The test launcher keeps your saved test definitions.
