# @lebedevna/xlsx

Simplifies building of excel-reports. Based on exceljs

## Examples

```ts
import { DataSheet, MetadataSheet, XlsxDocument, column } from "@lebedevna/xlsx";

const rows = [
  { product: "Keyboard", quantity: 2, price: 120 },
  { product: "Mouse", quantity: 4, price: 45 },
];

const metadata = new MetadataSheet("Об отчете", {
  entries: [
    ["Автор", "Иванов Иван Иванович"],
    ["Дата создания", new Date().toISOString()],
  ],
});

const data = new DataSheet("Данные", {
  elems: rows,
  headers: [
    column("product", "Товар", { width: 30 }),
    column("quantity", "Количество"),
    column("total", "Итого"),
  ],
  renderRow(row, ctx) {
    ctx.setValue("product", row.product);
    ctx.setValue("quantity", row.quantity);
    ctx.setValue("total", row.quantity * row.price);
  },
});

const document = new XlsxDocument([metadata, data]);
await document.renderToFile("report.xlsx");
```
