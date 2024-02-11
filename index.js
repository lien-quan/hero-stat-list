const axios = require("axios");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

async function scrapeDataAndGenerateExcel() {
  const baseUrl = "https://lienquan.garena.vn/tuong-chi-tiet/";
  const workbook = new ExcelJS.Workbook();

  for (let heroId = 1; heroId <= 130; heroId++) {
    console.info(`Scraping data for hero ID: ${heroId}`);
    const url = `${baseUrl}${heroId}`;
    let response = null;
    let $ = null;
    try {
      response = await axios.get(url);
      // Check if response status is not 200
      if (response.status !== 200) {
        console.log(
          `Skipping hero ID: ${heroId} due to non-200 response status.`
        );
        continue; // Skip to the next iteration of the loop
      }

      $ = cheerio.load(response.data);
      // Further processing...
    } catch (error) {
      console.error(
        `Error scraping data for hero ID: ${heroId}. Skipping to the next.`,
        error.message
      );
      continue; // Skip to the next iteration of the loop in case of error
    }

    // Extract the hero's name for the sheet name
    const heroName =
      $(".heroes-page .inner-page .skin-hero .title").text().trim() ||
      `Hero ${heroId}`;

    // Use the hero's name as the sheet name, with a fallback
    const sheet = workbook.addWorksheet(heroName.substring(0, 31)); // Excel sheet names must be <= 31 characters

    // Define columns
    sheet.columns = [
      { header: "Stat Name", key: "name", width: 30 },
      { header: "Initial Value", key: "initial", width: 15 },
      { header: "Increase Value", key: "increase", width: 15 },
    ];

    // Scrape data and populate the sheet
    $(".cont .col p").each((i, el) => {
      const statName = $(el).find("label").text().trim();
      const initialValue = $(el).find(".champion_stat").attr("data-original");
      const increaseValue = $(el).find(".champion_stat").attr("data-increase");

      // Add row to sheet
      sheet.addRow({
        name: statName,
        initial: initialValue,
        increase: increaseValue,
      });
    });
  }

  // Save the workbook to a file
  await workbook.xlsx.writeFile("AllHeroes.xlsx");
  console.log("Excel file has been created.");
}

scrapeDataAndGenerateExcel();
