import items from "./resultados.json" assert { type: "json" };
import ExcelJS from "exceljs";
import { groups } from "./utils/variables.js";

function getItemsCalculations() {
  const new_items = [...items];
  let calculations = [];

  new_items.map((item) => {
    const groupItemCode = item.CodigoGrupoItem;
    const groupInformations = groups.filter(
      (group) => group.code === groupItemCode
    );

    if (groupInformations.length > 0) {
      const salesPrice = item.PrecoDeVendaVigente;
      const costPrice = item.ValorDeCusto;
      const profitTarget = groupInformations[0].profit;

      // Prevenir Infinity ou NaN ao calcular a porcentagem
      let percentage;
      if (costPrice === 0) {
        percentage = "O custo está igual a 0"; // Evita o Infinity
      } else {
        percentage =
          calculateHowManyPercentTheSalesPriceIsAboveOrBelowTheCostPrice(
            costPrice,
            salesPrice
          );
      }

      // Tratar valores negativos
      if (percentage < 0) {
        percentage = `${Math.abs(percentage)}% (abaixo do custo)`;
      } else if (isNaN(percentage) || !isFinite(percentage)) {
        percentage = "Valor inválido"; // Tratando NaN ou Infinity
      }

      const isTheProfitBellowTheTarget = percentage < profitTarget;
      const isTheCostValueGreaterThanOrEqualTheSalesValue =
        costPrice >= salesPrice;
      const isSellByAC = groupInformations[0].AC === true ? "SIM" : "NÃO";
      const isHaveStock = item.SaldoEmEstoque > 0 ? "SIM" : "NÃO";

      let stringPercentage;
      let stringProfitTarget = `${profitTarget}%`;
      if (percentage >= 0) {
        stringPercentage = `${percentage}%`;
      }

      const new_item = {
        CodigoDoItem: item.CodigoDoItem,
        CodigoDeFabricacao: item.CodigoDeFabricacao.trim(),
        Nome: item.NomeDoItem.trim(),
        PrecoDeVenda: item.PrecoDeVendaVigente,
        ValorDeCusto: item.ValorDeCusto,
        PorcentagemAtual: percentage >= 0 ? stringPercentage : percentage,
        PorcentagemIdealDoGrupo: stringProfitTarget,
        ValorDeCustoAcimaDoDeVenda:
          isTheCostValueGreaterThanOrEqualTheSalesValue ? "SIM" : "NÃO",
        AbaixoDaMargem: isTheProfitBellowTheTarget ? "SIM" : "NÃO",
        VendidoPelaAC: isSellByAC,
        TemEstoque: isHaveStock
      };

      calculations.push(new_item);
    }
  });

  saveToExcel(calculations);
}

function calculateHowManyPercentTheSalesPriceIsAboveOrBelowTheCostPrice(
  custo,
  venda
) {
  const diffBetweenXAndY = venda - custo;
  const divideByLess = diffBetweenXAndY / custo;
  const result = 100 * divideByLess;
  const fixedResult = result.toFixed(0);
  return fixedResult;
}

async function saveToExcel(data) {
  const workbook = new ExcelJS.Workbook();

  // Filtrar os itens de acordo com as três condições
  const itemsWithStock = data.filter(
    (item) => item.TemEstoque === "SIM" && item.VendidoPelaAC === "NÃO"
  );
  const itemsWithStockAndSoldByAC = data.filter(
    (item) => item.TemEstoque === "SIM" && item.VendidoPelaAC === "SIM"
  );
  const itemsWithoutStock = data.filter((item) => item.TemEstoque === "NÃO");

  // Ordenar os itens de cada aba de acordo com a prioridade
  const sortedItemsWithStock = sortByPriority(itemsWithStock);
  const sortedItemsWithStockAndSoldByAC = sortByPriority(
    itemsWithStockAndSoldByAC
  );
  const sortedItemsWithoutStock = sortByPriority(itemsWithoutStock);

  // Criar a aba para itens com estoque
  const worksheetWithStock = workbook.addWorksheet("Tem Estoque");

  worksheetWithStock.columns = getWorksheetColumns();
  worksheetWithStock.views = [{ state: "frozen", ySplit: 1 }];

  sortedItemsWithStock.forEach((item) => addRowWithFormatting(worksheetWithStock, item));

  // Criar uma aba para itens com estoque e vendidos pela AC
  const worksheetWithStockAndSoldByAC = workbook.addWorksheet("Tem Estoque e Vendido pela AC");

  worksheetWithStockAndSoldByAC.columns = getWorksheetColumns();
  worksheetWithStockAndSoldByAC.views = [{ state: "frozen", ySplit: 1 }];

  sortedItemsWithStockAndSoldByAC.forEach((item) =>
    addRowWithFormatting(worksheetWithStockAndSoldByAC, item)
  );

  // Criar uma aba para itens sem estoque
  const worksheetWithoutStock = workbook.addWorksheet("Não Tem Estoque");

  worksheetWithoutStock.columns = getWorksheetColumns();
  worksheetWithoutStock.views = [{ state: "frozen", ySplit: 1 }];

  sortedItemsWithoutStock.forEach((item) => addRowWithFormatting(worksheetWithoutStock, item));

  // Salva o arquivo Excel
  await workbook.xlsx.writeFile("planilha-wr.xlsx");
  console.log("Dados filtrados salvos em planilha-wr.xlsx");
}

// Função para ordenar itens de acordo com a prioridade especificada
function sortByPriority(items) {
  return items.sort((a, b) => {
    const aPriority = a.ValorDeCustoAcimaDoDeVenda === "SIM" ? 2 : a.AbaixoDaMargem === "SIM" ? 1 : 0;
    const bPriority = b.ValorDeCustoAcimaDoDeVenda === "SIM" ? 2 : b.AbaixoDaMargem === "SIM" ? 1 : 0;

    return bPriority - aPriority;
  });
}

// Função para obter as colunas do worksheet
function getWorksheetColumns() {
  return [
    { header: "Código do Item", key: "CodigoDoItem", width: 25 },
    { header: "Código de Fabricação", key: "CodigoDeFabricacao", width: 20 },
    { header: "Nome", key: "Nome", width: 50 },
    { header: "Preço de Venda", key: "PrecoDeVenda", width: 25 },
    { header: "Valor de Custo", key: "ValorDeCusto", width: 25 },
    { header: "Porcentagem Atual", key: "PorcentagemAtual", width: 25 },
    { header: "Porcentagem Ideal", key: "PorcentagemIdealDoGrupo", width: 25 },
    { header: "Custo Maior Que a Venda", key: "ValorDeCustoAcimaDoDeVenda", width: 25 },
    { header: "Abaixo da Margem", key: "AbaixoDaMargem", width: 25 },
    { header: "Vendido Pela AC", key: "VendidoPelaAC", width: 25 },
    { header: "Tem Estoque", key: "TemEstoque", width: 20 },
  ];
}

// Função para adicionar uma linha com formatação de destaque
function addRowWithFormatting(worksheet, item) {
  const row = worksheet.addRow(item);

  row.getCell("PrecoDeVenda").numFmt = "R$ #,##0.00";
  row.getCell("ValorDeCusto").numFmt = "R$ #,##0.00";

  row.eachCell((cell, colNumber) => {
    if (colNumber !== 3) {
      cell.alignment = { vertical: "middle", horizontal: "center" };
    }
  });

  if (item.ValorDeCustoAcimaDoDeVenda === "SIM") {
    row.getCell("ValorDeCustoAcimaDoDeVenda").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFC7CE" }, // Vermelho claro
    };
  }

  if (item.AbaixoDaMargem === "SIM") {
    row.getCell("AbaixoDaMargem").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEB9C" }, // Amarelo claro
    };
  }

  if (item.TemEstoque === "SIM") {
    row.getCell("TemEstoque").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "C6EFCE" }, // Verde claro para destacar estoque
    };
  }
}

getItemsCalculations();
