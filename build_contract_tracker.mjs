import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "C:/Users/-/Documents/ChatGPT/проект1/outputs/contract-tracker";
await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Панель контроля");
const contracts = wb.worksheets.add("Договоры");
const milestones = wb.worksheets.add("Вехи");
const docs = wb.worksheets.add("Документы");
const lists = wb.worksheets.add("Справочники");

const navy = "#17365D", blue = "#D9EAF7", pale = "#F4F8FB", green = "#D9EAD3", yellow = "#FFF2CC", red = "#FCE4D6", gray = "#D9E1F2";
const title = (sheet, range, text) => {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[text]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: "#FFFFFF", size: 15 }, horizontalAlignment: "left", verticalAlignment: "center" };
};
const header = (sheet, range) => {
  sheet.getRange(range).format = { fill: blue, font: { bold: true, color: navy }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "all", style: "thin", color: "#B7C9D6" } };
};
const body = (sheet, range) => { sheet.getRange(range).format = { borders: { preset: "inside", style: "thin", color: "#D9E2F3" }, verticalAlignment: "center", wrapText: true }; };

lists.getRange("A1:D1").values = [["Статусы договора", "Типы вех", "Статусы вех", "Типы документов"]];
lists.getRange("A2:D7").values = [
  ["Черновик", "Запуск", "Не начато", "Договор"],
  ["В работе", "Промежуточная", "В работе", "Акт"],
  ["Приостановлен", "Приёмка", "Завершено", "Отчёт"],
  ["Завершён", "Оплата", "Просрочено", "Счёт"],
  ["Расторгнут", "Закрытие", "", "Переписка"],
  ["", "", "", "Иное"]
];
header(lists, "A1:D1"); body(lists, "A2:D7");
lists.getRange("A:D").format.columnWidth = 19;
lists.showGridLines = false;

title(contracts, "A1:N1", "Реестр договоров");
contracts.getRange("A3:N3").values = [["ID", "№ договора", "Наименование / предмет", "Контрагент", "Ответственный", "Дата начала", "Дата завершения", "Сумма", "Статус", "Дней до завершения", "Контроль", "Ссылка на договор", "Примечание", "Последнее обновление"]];
header(contracts, "A3:N3");
const contractRows = Array.from({length: 50}, () => [null, null, null, null, null, null, null, null, null, null, null, null, null, null]);
contracts.getRange("A4:N53").values = contractRows;
contracts.getRange("J4").formulas = [["=IF(G4=\"\",\"\",G4-TODAY())"]]; contracts.getRange("J4:J53").fillDown();
contracts.getRange("K4").formulas = [["=IF(I4=\"Завершён\",\"Завершён\",IF(G4=\"\",\"Нет даты\",IF(G4<TODAY(),\"Просрочен\",IF(G4-TODAY()<=30,\"Скоро\",\"В срок\"))))"]]; contracts.getRange("K4:K53").fillDown();
body(contracts, "A4:N53");
contracts.getRange("F4:G53").format.numberFormat = "yyyy-mm-dd";
contracts.getRange("H4:H53").format.numberFormat = "#,##0.00";
contracts.getRange("N4:N53").format.numberFormat = "yyyy-mm-dd";
contracts.getRange("I4:I53").dataValidation = { rule: { type: "list", formula1: "'Справочники'!$A$2:$A$6" } };
contracts.getRange("K4:K53").conditionalFormats.add("containsText", { text: "Просрочен", format: { fill: red, font: { color: "#9C0006", bold: true } } });
contracts.getRange("K4:K53").conditionalFormats.add("containsText", { text: "Скоро", format: { fill: yellow, font: { color: "#7F6000", bold: true } } });
contracts.getRange("K4:K53").conditionalFormats.add("containsText", { text: "В срок", format: { fill: green, font: { color: "#006100" } } });
contracts.getRange("A:N").format.autofitColumns();
contracts.getRange("C:C").format.columnWidth = 30; contracts.getRange("M:M").format.columnWidth = 28;
contracts.getRange("A3:N53").format.rowHeight = 22; contracts.getRange("A3:N3").format.rowHeight = 40;
contracts.freezePanes.freezeRows(3); contracts.showGridLines = false;

title(milestones, "A1:L1", "Вехи и контроль этапов");
milestones.getRange("A3:L3").values = [["ID", "№ договора", "Наименование этапа", "Тип вехи", "Плановая дата", "Фактическая дата", "Дней до срока", "Статус контроля", "Подтверждающий документ", "Ссылка на документ", "Ответственный", "Комментарий"]];
header(milestones, "A3:L3");
milestones.getRange("A4:L103").values = Array.from({length: 100}, () => [null, null, null, null, null, null, null, null, null, null, null, null]);
milestones.getRange("G4").formulas = [["=IF(F4<>\"\",0,IF(E4=\"\",\"\",E4-TODAY()))"]]; milestones.getRange("G4:G103").fillDown();
milestones.getRange("H4").formulas = [["=IF(F4<>\"\",\"Завершено\",IF(E4=\"\",\"Нет даты\",IF(E4<TODAY(),\"Просрочено\",IF(E4-TODAY()<=14,\"Скоро\",\"В срок\"))))"]]; milestones.getRange("H4:H103").fillDown();
body(milestones, "A4:L103");
milestones.getRange("D4:D103").dataValidation = { rule: { type: "list", formula1: "'Справочники'!$B$2:$B$6" } };
milestones.getRange("E4:F103").format.numberFormat = "yyyy-mm-dd";
milestones.getRange("H4:H103").conditionalFormats.add("containsText", { text: "Просрочено", format: { fill: red, font: { color: "#9C0006", bold: true } } });
milestones.getRange("H4:H103").conditionalFormats.add("containsText", { text: "Скоро", format: { fill: yellow, font: { color: "#7F6000", bold: true } } });
milestones.getRange("H4:H103").conditionalFormats.add("containsText", { text: "Завершено", format: { fill: green, font: { color: "#006100" } } });
milestones.getRange("A:L").format.autofitColumns(); milestones.getRange("C:C").format.columnWidth = 28; milestones.getRange("L:L").format.columnWidth = 28;
milestones.getRange("A3:L103").format.rowHeight = 22; milestones.getRange("A3:L3").format.rowHeight = 38;
milestones.freezePanes.freezeRows(3); milestones.showGridLines = false;

title(docs, "A1:J1", "Реестр документов по договору и этапам");
docs.getRange("A3:J3").values = [["ID", "№ договора", "Этап / веха", "Тип документа", "Название документа", "Дата документа", "Номер", "Ссылка / путь к файлу", "Проверен", "Комментарий"]];
header(docs, "A3:J3");
docs.getRange("A4:J103").values = Array.from({length: 100}, () => [null, null, null, null, null, null, null, null, null, null]);
body(docs, "A4:J103"); docs.getRange("D4:D103").dataValidation = { rule: { type: "list", formula1: "'Справочники'!$D$2:$D$7" } };
docs.getRange("I4:I103").dataValidation = { rule: { type: "list", values: ["Да", "Нет"] } };
docs.getRange("F4:F103").format.numberFormat = "yyyy-mm-dd";
docs.getRange("A:J").format.autofitColumns(); docs.getRange("E:E").format.columnWidth = 28; docs.getRange("H:H").format.columnWidth = 32; docs.getRange("J:J").format.columnWidth = 25;
docs.getRange("A3:J103").format.rowHeight = 22; docs.getRange("A3:J3").format.rowHeight = 38;
docs.freezePanes.freezeRows(3); docs.showGridLines = false;

title(dashboard, "A1:H1", "Контроль исполнения договоров");
dashboard.getRange("A3:B7").values = [["Показатель", "Значение"], ["Всего договоров", null], ["Договоры в работе", null], ["Просроченные договоры", null], ["Вехи просрочены", null]];
dashboard.getRange("D3:E7").values = [["Ближайшие действия", "Количество"], ["Вехи в ближайшие 14 дней", null], ["Договоры завершатся в ближайшие 30 дней", null], ["Непроверенные документы", null], ["Завершённые вехи", null]];
dashboard.getRange("B4").formulas = [["=COUNTIF('Договоры'!$B$4:$B$53,\"?*\")"]];
dashboard.getRange("B5").formulas = [["=COUNTIF('Договоры'!$I$4:$I$53,\"В работе\")"]];
dashboard.getRange("B6").formulas = [["=COUNTIF('Договоры'!$K$4:$K$53,\"Просрочен\")"]];
dashboard.getRange("B7").formulas = [["=COUNTIF('Вехи'!$H$4:$H$103,\"Просрочено\")"]];
dashboard.getRange("E4").formulas = [["=COUNTIF('Вехи'!$H$4:$H$103,\"Скоро\")"]];
dashboard.getRange("E5").formulas = [["=COUNTIF('Договоры'!$K$4:$K$53,\"Скоро\")"]];
dashboard.getRange("E6").formulas = [["=COUNTIFS('Документы'!$B$4:$B$103,\"?*\",'Документы'!$I$4:$I$103,\"Нет\")"]];
dashboard.getRange("E7").formulas = [["=COUNTIF('Вехи'!$H$4:$H$103,\"Завершено\")"]];
header(dashboard, "A3:B3"); header(dashboard, "D3:E3");
dashboard.getRange("A4:B7").format = { fill: pale, borders: { preset: "all", style: "thin", color: "#B7C9D6" } };
dashboard.getRange("D4:E7").format = { fill: pale, borders: { preset: "all", style: "thin", color: "#B7C9D6" } };
dashboard.getRange("B4:B7").format = { fill: gray, font: { bold: true, color: navy, size: 14 }, horizontalAlignment: "center" };
dashboard.getRange("E4:E7").format = { fill: gray, font: { bold: true, color: navy, size: 14 }, horizontalAlignment: "center" };
dashboard.getRange("A10:H13").merge(); dashboard.getRange("A10:H13").values = [["Как работать с реестром:\n1. Заполните карточку договора на листе «Договоры».\n2. Добавьте все контрольные точки на листе «Вехи».\n3. Вносите акты, отчёты и иные подтверждения на листе «Документы».\n4. Контрольные статусы и панель обновляются автоматически."]];
dashboard.getRange("A10:H13").format = { fill: "#EAF2F8", font: { color: navy }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: "#B7C9D6" } };
dashboard.getRange("A:B").format.columnWidth = 26; dashboard.getRange("D:D").format.columnWidth = 34; dashboard.getRange("E:E").format.columnWidth = 24; dashboard.getRange("A1:H1").format.rowHeight = 28; dashboard.getRange("A10:H13").format.rowHeight = 28;
dashboard.showGridLines = false;

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outputDir}/monitoring_dogovorov.xlsx`);

for (const [sheetName, fileName] of [["Панель контроля", "dashboard.png"], ["Договоры", "contracts.png"], ["Вехи", "milestones.png"], ["Документы", "docs.png"]]) {
  const image = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await image.arrayBuffer()));
}

console.log(JSON.stringify({ output: `${outputDir}/monitoring_dogovorov.xlsx` }));
