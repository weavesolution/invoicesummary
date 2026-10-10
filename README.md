# Invoice Summariser

Static web page (GitHub Pages) that reads tax-invoice PDFs / ZIPs in the browser and produces an item-wise summary plus a formatted Excel file. Nothing is uploaded to a server.

## Publish on GitHub Pages
1. Create a repository and upload everything in this folder (keep the `js/` and `lib/` folders).
2. Settings > Pages > Deploy from a branch > `main` / root.
3. Open `https://<your-username>.github.io/<repo>/`.

## Output columns
Invoice No, GST number, HSN Code, Item Name, Qty, Packaging, Discount Allocated, Total Amount, CGST, SGST, IGST, Total Incl. GST.

- GST number is the customer's GSTIN; blank becomes `non-registered`.
- Whole-bill discount goes to item 1; GST is split across items by taxable value.
- On the Global Dates bill, the line printed "Add : GST 2.50%" is read as SGST.
- A second sheet, "Invoice Check", compares every invoice with the total printed on the bill.

## Adding a new bill design
Open `js/core.js`, copy the `global-dates-v1` parser, change `detect()` and `parse()`, and register it. Everything else is shared.
