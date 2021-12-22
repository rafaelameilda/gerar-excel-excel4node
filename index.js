const express = require("express");
const app = express();
const port = 3000;

const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("Worksheet Name");

const data = [
  {
    name: "Teste",
    email: "teste@gmail.com",
    cellphone: "1234567890",
  },
  {
    name: "Pessoa 2",
    email: "pessoa@gmail.com",
    cellphone: "1234567899",
  },
];

const headingColumnNames = ["Nome", "Email", "Celular"];

let headingColumnIndex = 1; //diz que começará na primeira linha
headingColumnNames.forEach((heading) => {
  //passa por todos itens do array
  // cria uma célula do tipo string para cada título
  ws.cell(1, headingColumnIndex++).string(heading);
});

let rowIndex = 2;
data.forEach((record) => {
  let columnIndex = 1;
  Object.keys(record).forEach((columnName) => {
    ws.cell(rowIndex, columnIndex++).string(record[columnName]);
  });
  rowIndex++;
});

app.get("/", (req, res) => {
  wb.write("ArquivoExcel.xlsx", res);
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
