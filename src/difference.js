import { groups } from "./utils/variables.js";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

// Função para carregar dados diretamente do arquivo JSON
function loadDataFromJson() {
  const filePath = path.join(process.cwd(), "src", "resultados.json");
  if (!fs.existsSync(filePath)) {
    throw new Error(`Arquivo ${filePath} não encontrado.`);
  }
  const data = fs.readFileSync(filePath, "utf-8");
  return JSON.parse(data);
}

// Função para criar o arquivo Excel com formatação
async function createExcelFile(data) {
  const workbook = new ExcelJS.Workbook();

  // Agrupar itens por grupo
  const groupedItems = data.reduce((acc, item) => {
    if (!acc[item.item_group]) acc[item.item_group] = [];
    acc[item.item_group].push(item);
    return acc;
  }, {});

  // Criar uma aba para cada grupo de itens
  for (const [group, items] of Object.entries(groupedItems)) {
    const sanitizedGroup = group.replace(/[*/?:\\[\]]/g, ""); // Remover caracteres inválidos do nome
    const worksheet = workbook.addWorksheet(sanitizedGroup);

    // Adicionar cabeçalho com largura personalizada e centralização
    worksheet.columns = [
      { header: "Código de Fabricação", key: "fabrication_code", width: 30 },
      { header: "Nome do Item", key: "name", width: 60 }, // Dobrando o tamanho da coluna
      { header: "Preço Antigo", key: "old_price", width: 25 },
      { header: "Novo Preço", key: "new_price", width: 25 },
      { header: "Porcentagem de aumento", key: "percentage_increase", width: 30 },
    ];

    // Centralizar cabeçalhos
    worksheet.columns.forEach((column) => {
      column.alignment = { vertical: "middle", horizontal: "center" };
    });

    // Adicionar dados e centralizar todas as colunas
    items.forEach((item) => {
      // Lógica para calcular a porcentagem de aumento
      let percentage_increase;
      if (item.old_price === 0 && item.new_price > 0) {
        percentage_increase = "Valor de venda estava 0"; // Caso especial para preços antigos igual a 0
      } else {
        percentage_increase = Math.round(((item.new_price - item.old_price) / item.old_price) * 100);
      }

      // Ignorar itens com porcentagem negativa
      if (percentage_increase !== "Valor de venda estava 0" && percentage_increase < 0) return;

      const row = worksheet.addRow({
        fabrication_code: item.fabrication_code,
        name: item.name,
        old_price: `R$ ${item.old_price.toFixed(2)}`, // Formatar preço antigo
        new_price: `R$ ${item.new_price.toFixed(2)}`, // Formatar novo preço
        percentage_increase: percentage_increase === "Valor de venda estava 0" 
          ? percentage_increase 
          : `${percentage_increase}%`, // Adiciona a porcentagem ou a mensagem
      });

      // Centralizar células
      row.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });
    });

    // Congelar a linha de cabeçalho
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
  }

  // Salvar o arquivo Excel
  const filePath = path.join(process.cwd(), "src", "resultados.xlsx");
  await workbook.xlsx.writeFile(filePath);
  console.log(`Arquivo Excel salvo em ${filePath}`);
}

(async () => {
  try {
    const data = loadDataFromJson();
    await createExcelFile(data);
  } catch (error) {
    console.error("Erro no fluxo principal:", error.message);
  }
})();
