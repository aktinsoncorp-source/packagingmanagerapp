MISSING PRODUCT TEST - Packing List Manager V2.0.0

Close any open V2 Packing List Manager window, then double-click
Start Missing Product Test.cmd in this folder.
It opens the Python app with independent test names and settings in AppData.
The normal application's ProductNames.json and Settings.json are not used.

The test date is 10 October 2026 (20261010).
If choosing the folder manually, choose this Test Missing Product folder,
set the packing date to 10 October 2026, and click Refresh.
Use the test launcher first to keep your normal definitions separate.

There are two synthetic FLASH orders:

1. Missing SKU and standard names
Order ID: TESTNOSKU20261010A
Ecommerce Name: TEST PRODUCT - MOONLIGHT VANILLA
Ecommerce Variation: 50 mL
Quantity: 1
Seller SKU: blank
Prices: unknown (blank)

2. SKU present, missing standard names
Order ID: TESTWITHSKU20261010B
Ecommerce Name: TEST PRODUCT - AURORA PEARL
Ecommerce Variation: 50 mL
Quantity: 1
Seller SKU: TEST-AURORA-50ML
Product Name and Size: blank until you define them in the app
Prices: unknown (blank)

These products intentionally do not exist in the current product catalog.
They are for testing only. Do not ship them or use them as real customer orders.

Try Later/ignoring the names prompt, then Generate Packing Lists. The test
output should retain the ecommerce name and variation. The reminder should
continue to direct attention to the missing SKU and names.

Then click Define now or Review. Editing opens unlocked for reviewed products;
undefined rows remain unlocked until Product Name and Size are complete.
Enter a Product Name and Size. Save names, then generate again if desired.
AURORA PEARL should stop needing review after both fields are saved.
A missing-SKU reminder may remain until its ecommerce SKU has been corrected.
Test output is saved beneath 20261010/000 Output.

Keep the existing source PDFs; use completed-document/reprint controls to
repeat a generated order. The test launcher keeps your saved test definitions.
