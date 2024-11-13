import { items_by_ac_before, groups } from "./utils/variables.js";
import axios from "axios";
import puppeteer from "puppeteer";
import cliProgress from "cli-progress";
import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

async function launchBrowser() {
  try {
    console.log("Iniciando o Puppeteer...");
    const browser = await puppeteer.launch({
      executablePath:
        "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    });
    const page = await browser.newPage();
    console.log("Puppeteer iniciado.");
    return { browser, page };
  } catch (error) {
    console.error("Erro ao iniciar o Puppeteer:", error.message);
    throw error;
  }
}

async function loginToMovere(page) {
  try {
    console.log("Navegando até a página de login...");
    await page.goto(
      "https://bsb.moveresoftware.com/danielautocenter/profile/logon.aspx?",
      { waitUntil: "networkidle2" }
    );
    console.log("Página de login carregada.");

    if (
      !(await page.$("#ctl00_Corpo_edtUsername")) ||
      !(await page.$("#ctl00_Corpo_edtPassword"))
    ) {
      throw new Error("Campos de login não encontrados.");
    }

    console.log("Preenchendo o formulário de login...");
    await page.type("#ctl00_Corpo_edtUsername", "weslley.santos");
    await page.type("#ctl00_Corpo_edtPassword", "100120");

    console.log("Enviando formulário...");
    await Promise.all([
      page.click("#ctl00_Corpo_btnConnect"),
      page.waitForNavigation(),
    ]);
    console.log("Login realizado com sucesso.");

    const cookies = await page.cookies();
    await page.browser().close();
    console.log("Cookies obtidos e navegador fechado.");
    return cookies;
  } catch (error) {
    console.error("Erro no processo de login:", error.message);
    throw error;
  }
}

function formatCookies(cookies) {
  try {
    console.log("Cookies convertidos para o formato do Axios.");
    return cookies.map((cookie) => `${cookie.name}=${cookie.value}`).join("; ");
  } catch (error) {
    console.error("Erro ao formatar cookies:", error.message);
    throw error;
  }
}

const getDifferenceBetweenBeforeAndAfterPriceChanges = async (
  cookieHeader,
  maxRetries = 1
) => {
  const new_items_by_ac_before = [...items_by_ac_before];
  const progressBar = new cliProgress.SingleBar(
    {
      format: "Progresso | {bar} | {percentage}% | Requisição {value}/{total}",
    },
    cliProgress.Presets.shades_classic
  );

  progressBar.start(new_items_by_ac_before.length, 0);

  let results = [];
  let failedRequests = [];
  for (const item of new_items_by_ac_before) {
    const { item_code, name, fabrication_code, sale_price } = item;
    let attempt = 0;
    let success = false;

    while (attempt < maxRetries && !success) {
      try {
        attempt++;
        const response = await axios.get(
          `https://bsb.moveresoftware.com/danielautocenter/rot.mvc/R11/BuscarItens?termo=${fabrication_code.trim()}&codigoDeposito=1&codigoDaLoja=1&tipoLocalEntrega=2&tabelaDePreco=1`,
          {
            headers: { Cookie: cookieHeader },
            timeout: 10000,
          }
        );

        if (Array.isArray(response.data)) {
          const matchedItem = response.data.find(
            (apiItem) => apiItem.CodigoDoItem === item_code
          );
          if (matchedItem) {
            const new_price = matchedItem.PrecoDeVendaVigente;
            const old_price = sale_price;

            // Ignora itens cujo preço não mudou
            if (old_price === new_price) {
              success = true;
              continue; // Pule para o próximo item no loop
            }

            // Calcular aumento percentual
            let percentage_increase;
            if (old_price === 0) {
              percentage_increase = "Novo preço indefinido";
            } else {
              percentage_increase = ((new_price - old_price) / old_price) * 100;
            }

            // Localizar o item_group usando groups
            const item_group =
              groups.find((group) => group.code === matchedItem.CodigoGrupoItem)
                ?.name || "Grupo não encontrado";

            const compared_item = {
              fabrication_code,
              name,
              old_price,
              new_price,
              percentage_increase,
              item_group,
            };
            results.push(compared_item);
            success = true;
          } else {
            console.error(
              `Item com code "${item_code}" e name "${name}" não encontrado na resposta.`
            );
            throw new Error(
              `Item com code "${item_code}" e name "${name}" não encontrado.`
            );
          }
        } else {
          throw new Error("Resposta da API não contém um array.");
        }
      } catch (error) {
        console.error(
          `Erro na requisição para o code "${item_code}" e name "${name}" (tentativa ${attempt}/${maxRetries}):`,
          error.message
        );

        if (attempt === maxRetries) {
          failedRequests.push({ item_code, name, error: error.message });
        }

        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }

    progressBar.increment();
  }

  progressBar.stop();

  // Ordena os resultados pelo percentual de aumento, ignorando os casos "Novo preço indefinido"
  results.sort((a, b) => {
    if (
      typeof a.percentage_increase === "number" &&
      typeof b.percentage_increase === "number"
    ) {
      return b.percentage_increase - a.percentage_increase;
    }
    return typeof a.percentage_increase === "number" ? -1 : 1;
  });

  // Salvar os resultados em um arquivo Excel em vez de JSON
  await saveResultsToExcel(results);
};

// Função para salvar os resultados no arquivo Excel
const saveResultsToExcel = async (results) => {
  try {
    const workbook = new ExcelJS.Workbook();

    // Agrupar os itens por grupo
    const groupedResults = results.reduce((acc, item) => {
      const { item_group } = item;
      if (!acc[item_group]) acc[item_group] = [];
      acc[item_group].push(item);
      return acc;
    }, {});

    // Para cada grupo, cria uma aba
    for (const [groupName, items] of Object.entries(groupedResults)) {
      // Limpar o nome da aba para remover caracteres inválidos
      const sanitizedGroupName = groupName.replace(/[*?:/\\[\]]/g, "_");
      const worksheet = workbook.addWorksheet(sanitizedGroupName);

      // Definir cabeçalho com nomes personalizados
      worksheet.columns = [
        { header: "Código de Fabricação", key: "fabrication_code", width: 20 },
        { header: "Nome do Item", key: "name", width: 30 },
        { header: "Preço Antigo", key: "old_price", width: 15 },
        { header: "Novo Preço", key: "new_price", width: 15 },
        { header: "Porcentagem de Aumento", key: "percentage_increase", width: 20 },
      ];

      // Adiciona os dados
      items.forEach((item) => worksheet.addRow(item));

      // Congelar a linha de cabeçalho para sempre ficar visível
      worksheet.views = [{ state: "frozen", ySplit: 1 }];
    }

    // Definir o caminho e salvar o arquivo
    const filePath = path.join(process.cwd(), "src", "resultados.xlsx");
    await workbook.xlsx.writeFile(filePath);
    console.log(`Arquivo Excel salvo em: ${filePath}`);
  } catch (error) {
    console.error("Erro ao salvar o arquivo Excel:", error.message);
    throw error;
  }
};


(async () => {
  try {
    const { browser, page } = await launchBrowser();
    const cookies = await loginToMovere(page);
    const cookieHeader = formatCookies(cookies);
    await getDifferenceBetweenBeforeAndAfterPriceChanges(cookieHeader);
  } catch (error) {
    console.error("Erro no fluxo principal:", error.message);
  }
})();
