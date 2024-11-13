import axios from "axios";
import puppeteer from "puppeteer";
import cliProgress from "cli-progress";
import fs from "fs";
import path from "path";
import { group_items } from "./utils/variables.js";

async function launchBrowser() {
  try {
    console.log("Iniciando o Puppeteer...");
    const browser = await puppeteer.launch({
      executablePath:
        "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
      headless: false,
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

async function fetchData(id_list, cookieHeader, maxRetries = 1) {
  console.log("Realizando requisição à API...");

  const progressBar = new cliProgress.SingleBar(
    {
      format: "Progresso | {bar} | {percentage}% | Requisição {value}/{total}",
    },
    cliProgress.Presets.shades_classic
  );

  progressBar.start(id_list.length, 0);

  let results = [];
  let failedRequests = [];

  for (const item of id_list) {
    const { code, name, fabrication } = item;
    let attempt = 0;
    let success = false;

    while (attempt < maxRetries && !success) {
      try {
        attempt++;
        const response = await axios.get(
          `https://bsb.moveresoftware.com/danielautocenter/rot.mvc/R11/BuscarItens?termo=${fabrication.trim()}&codigoDeposito=1&codigoDaLoja=1&tipoLocalEntrega=2&tabelaDePreco=1`,
          {
            headers: { Cookie: cookieHeader },
            timeout: 10000,
          }
        );

        if (Array.isArray(response.data)) {
          const matchedItem = response.data.find(
            (apiItem) => apiItem.CodigoDoItem === code
          );
          if (matchedItem) {
            results.push(matchedItem);
            success = true;
          } else {
            console.error(
              `Item com code "${code}" e name "${name}" não encontrado na resposta.`
            );
            throw new Error(
              `Item com code "${code}" e name "${name}" não encontrado.`
            );
          }
        } else {
          throw new Error("Resposta da API não contém um array.");
        }
      } catch (error) {
        console.error(
          `Erro na requisição para o code "${code}" e name "${name}" (tentativa ${attempt}/${maxRetries}):`,
          error.message
        );

        if (attempt === maxRetries) {
          failedRequests.push({ code, name, error: error.message });
        }

        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }

    progressBar.increment();
  }

  progressBar.stop();

  const filePath = path.join(process.cwd(), "src", "resultados.json");
  const failedPath = path.join(process.cwd(), "src", "falhas_requisicoes.json");

  try {
    if (results.length > 0) {
      fs.writeFileSync(filePath, JSON.stringify(results, null, 2), "utf-8");
      console.log(`Dados salvos em ${filePath}`);
    } else {
      console.warn("Nenhum dado foi adicionado ao arquivo `resultados.json`.");
    }

    if (failedRequests.length > 0) {
      fs.writeFileSync(
        failedPath,
        JSON.stringify(failedRequests, null, 2),
        "utf-8"
      );
      console.log(`Log de falhas salvo em ${failedPath}`);
    } else {
      console.log("Nenhuma falha de requisição registrada.");
    }
  } catch (error) {
    console.error("Erro ao salvar arquivos JSON:", error.message);
    throw error;
  }
}

(async () => {
  try {
    const { browser, page } = await launchBrowser();
    const cookies = await loginToMovere(page);
    const cookieHeader = formatCookies(cookies);
    await fetchData(group_items, cookieHeader);
  } catch (error) {
    console.error("Erro no fluxo principal:", error.message);
  }
})();
