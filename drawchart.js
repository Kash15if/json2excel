const { IData, XlsxGenerator } = require("office-chart");

const gen = new XlsxGenerator();

gen.createWorkbook();

const sheet1 = gen.createWorksheet("sheet1");

const header = ["h", "b", "c", "d"];
const row1 = ["label1", 2, 3, 4];
const row2 = ["label2", 5, 6, 7];

sheet2.gen.generate(__dirname + "/test", "file");
