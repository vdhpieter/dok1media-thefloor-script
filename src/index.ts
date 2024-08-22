import fs from "node:fs";
import neatCsv from "neat-csv";
import xlsx from "xlsx";
import { decode } from "html-entities";

xlsx.set_fs(fs);

const questionsDir = "data/Questions";
const PhotosDir = "data/Photos";
const audioDir = "data/Audio";

const categoryFolders = fs.readdirSync(questionsDir);

let result =
  "\uFEFFVak,Recordnummer,Hoofdcategorie,Subcategorie,Vraag presentatie,Vraag graphics,Subcategorie presentatie,Volgorde,Antwoord graphics,Alle antwoorden,Antwoord dubbele laag,Vraag,Antwoord A,Antwoord B,Antwoord C,Meerkeuzeantwoord,Associatie 1,Associatie 2,Associatie 3,Associatie 4,Vraagsoort\n";

type Question = {
  "\uFEFFNr Vlak Vloer": string;
  Hoofdcat: string;
  Subcat: string;
  "Soort Vraag": string;
  "Volgorde vraag": string;
  "Presentatie Vraag": string;
  "Grafiek Vraag": string;
  Antwoord: string;
  "Extra Antwoorden / Shadow": string;
  "Associatie 1": string;
  "Associatie 2": string;
  "Associatie 3": string;
  "Associatie 4": string;
  Geluid: string;
  "Meerkeuze A": string;
  "Meerkeuze B": string;
  "Meerkeuze C": string;
};

for (const categoryFolder of categoryFolders) {
  if (categoryFolder === ".DS_Store") {
    continue;
  }
  const subCategoryFiles = fs.readdirSync(`${questionsDir}/${categoryFolder}`);

  for (const questionFile of subCategoryFiles) {
    if (questionFile === ".DS_Store") {
      continue;
    }
    if (questionFile.startsWith("~$")) {
      continue;
    }
    if (questionFile.endsWith(".csv")) {
      continue;
    }

    const workbookName = `${questionsDir}/${categoryFolder}/${questionFile}`;
    const workBook = xlsx.readFile(workbookName, {});

    if (
      String(workBook.Sheets[workBook.SheetNames[0]]["F1"].v).match(
        /Volgorde.*\n/
      )
    ) {
      console.log("rewriting volgorde heading");
      xlsx.utils.sheet_add_aoa(
        workBook.Sheets[workBook.SheetNames[0]],
        [["Volgorde vraag"]],
        { origin: "F1" }
      );
    }

    if (
      String(workBook.Sheets[workBook.SheetNames[0]]["E1"].v).match(/Soort.*\n/)
    ) {
      console.log("rewriting soort heading");
      xlsx.utils.sheet_add_aoa(
        workBook.Sheets[workBook.SheetNames[0]],
        [["Soort Vraag"]],
        { origin: "E1" }
      );
    }

    const csvName = workbookName.replace("xlsx", "csv");
    xlsx.writeFile(workBook, csvName, { bookType: "csv" });
    const rawData = fs.createReadStream(csvName, { encoding: "utf-8" });
    const questions = await neatCsv<Question>(rawData, { separator: "," });

    console.log(`working on ${workbookName}`);

    if (questions[0]["Soort Vraag"].toLowerCase() === "foto") {
      const vak = questions[0]["\uFEFFNr Vlak Vloer"];
      const hoofdcategorie = questions[0].Hoofdcat;
      const subcategorie = questions[0].Subcat;
      const vraagPresentatie = questions[0]["Presentatie Vraag"];
      const vraagGraphics = questions[0]["Grafiek Vraag"];

      for (const question of questions) {
        if (question["Volgorde vraag"] === "") {
          break;
        }
        result = result.concat(
          `${vak},${vak},${hoofdcategorie},${subcategorie},`
        );
        result = result.concat(
          `"${vraagPresentatie}",${vraagGraphics},,${question["Volgorde vraag"]},"${question.Antwoord}","${question["Extra Antwoorden / Shadow"]}",`
        );
        result = result.concat(`,,,,,,,,,,Foto's\n`);

        const subCategoryPhotoDir = questionFile.split(" -")[1].trim();

        const photosForCategoryDir = `${categoryFolder}/${subCategoryPhotoDir}`;
        const inputDir = `${PhotosDir}/${photosForCategoryDir}/EXPORT`;
        const allPhotoFiles = fs.readdirSync(inputDir);
        const photoFiles = allPhotoFiles.filter((photo) => {
          return (
            photo.includes(`_${question["Volgorde vraag"]}.`) ||
            photo.includes(`_${question["Volgorde vraag"]} `) ||
            photo.startsWith(`${question["Volgorde vraag"]}.`) ||
            photo.startsWith(`${question["Volgorde vraag"]}_`) ||
            photo.startsWith(`${question["Volgorde vraag"].padStart(2, "0")}.`)
          );
        });
        const outputDir = `output/${photosForCategoryDir}`;
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }
        if (photoFiles.length === 1) {
          fs.copyFileSync(
            `${inputDir}/${photoFiles[0]}`,
            `${outputDir}/${vak}_V${question["Volgorde vraag"]}.jpg`
          );
        } else if (photoFiles.length === 2) {
          const indexOfRevealPhoto = photoFiles.findIndex((photo) => {
            return photo.toLowerCase().includes("reveal");
          });

          const indexOfQuestionPhoto = indexOfRevealPhoto === 0 ? 1 : 0;
          if (photoFiles[indexOfQuestionPhoto].includes("reveal")) {
            throw `2 reveal photos for ${photosForCategoryDir} ${question["Volgorde vraag"]}`;
          }

          fs.copyFileSync(
            `${inputDir}/${photoFiles[indexOfQuestionPhoto]}`,
            `${outputDir}/${vak}_V${question["Volgorde vraag"]}.jpg`
          );
          fs.copyFileSync(
            `${inputDir}/${photoFiles[indexOfRevealPhoto]}`,
            `${outputDir}/${vak}_R${question["Volgorde vraag"]}.jpg`
          );
        } else {
          throw `${photoFiles.length} photos for ${photosForCategoryDir} ${
            question["Volgorde vraag"]
          }: ${JSON.stringify(photoFiles)}`;
        }
      }
    } else if (questions[0]["Soort Vraag"].toLowerCase() === "geluid") {
      const vak = questions[0]["\uFEFFNr Vlak Vloer"];
      const hoofdcategorie = questions[0].Hoofdcat;
      const subcategorie = questions[0].Subcat;
      const vraagPresentatie = questions[0]["Presentatie Vraag"];
      const vraagGraphics = questions[0]["Grafiek Vraag"];

      for (const question of questions) {
        if (question["Volgorde vraag"] === "") {
          break;
        }
        result = result.concat(
          `${vak},${vak},${hoofdcategorie},${subcategorie},`
        );
        result = result.concat(
          `"${vraagPresentatie}",${vraagGraphics},,${question["Volgorde vraag"]},"${question.Antwoord}","${question["Extra Antwoorden / Shadow"]}","${question.Geluid}",`
        );
        result = result.concat(`,,,,,,,,,Audio\n`);

        const subCategoryAudioDir = questionFile.split(" -")[1].trim();

        const audioForCategoryDir = `${categoryFolder}/${subCategoryAudioDir}`;
        const inputDir = `${audioDir}/${audioForCategoryDir}`;
        const allAudioFiles = fs.readdirSync(inputDir);
        const audioFiles = allAudioFiles.filter((audio) => {
          return (
            audio.startsWith(`${question["Volgorde vraag"]}.`) ||
            audio.startsWith(`${question["Volgorde vraag"].padStart(2, "0")}.`)
          );
        });
        const outputDir = `output/${audioForCategoryDir}`;
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir, { recursive: true });
        }
        if (audioFiles.length === 1) {
          fs.copyFileSync(
            `${inputDir}/${audioFiles[0]}`,
            `${outputDir}/${vak}_V${question["Volgorde vraag"]}.wav`
          );
        } else {
          throw `${audioFiles.length} sounds for ${audioForCategoryDir} ${
            question["Volgorde vraag"]
          }: ${JSON.stringify(audioFiles)}`;
        }
      }
    } else if (questions[0]["Soort Vraag"].toLowerCase() === "meerkeuze") {
      const vak = questions[0]["\uFEFFNr Vlak Vloer"];
      const hoofdcategorie = questions[0].Hoofdcat;
      const subcategorie = questions[0].Subcat;

      for (const question of questions) {
        if (question["Volgorde vraag"] === "") {
          break;
        }
        result = result.concat(
          `${vak},${vak},${hoofdcategorie},${subcategorie},`
        );
        result = result.concat(
          `,,,${question["Volgorde vraag"]},,,,"${question[
            "Presentatie Vraag"
            //@ts-ignore - we use a recent enough version
          ].replaceAll('"', '""')}",${question["Meerkeuze A"]},${
            question["Meerkeuze B"]
          },${question["Meerkeuze C"]},${question.Antwoord.charAt(0)},`
        );
        result = result.concat(`,,,,Meerkeuze\n`);
      }
    } else if (questions[0]["Soort Vraag"].toLowerCase() === "associatie") {
      const vak = questions[0]["\uFEFFNr Vlak Vloer"];
      const hoofdcategorie = questions[0].Hoofdcat;
      const subcategorie = questions[0].Subcat;
      const vraagPresentatie = questions[0]["Presentatie Vraag"];
      const vraagGraphics = questions[0]["Grafiek Vraag"];

      for (const question of questions) {
        if (question["Volgorde vraag"] === "") {
          break;
        }
        result = result.concat(
          `${vak},${vak},${hoofdcategorie},${subcategorie},`
        );
        result = result.concat(
          `"${vraagPresentatie}","${vraagGraphics}",,${
            question["Volgorde vraag"]
          },"${question.Antwoord}","${
            question["Extra Antwoorden / Shadow"]
          }",,,,,,,"${question["Associatie 1"]}","${
            question["Associatie 2"]
          }","${question["Associatie 3"]}","${question[
            "Associatie 4"
            //@ts-ignore - we use a recent enough version
          ].replaceAll('"', '""')}",`
        );
        result = result.concat(`Associaties\n`);
      }
    } else {
      throw `Unknown question type in ${workbookName}: ${questions[0]["Soort Vraag"]}`;
    }
  }
}
const outputFile = "output/thefloor-belgie-S01-vragen.csv";

fs.writeFileSync(outputFile, decode(result));
const resultWorkbook = xlsx.readFile(outputFile);
xlsx.utils.sheet_add_aoa(
  resultWorkbook.Sheets[resultWorkbook.SheetNames[0]],
  [["Vak"]],
  { origin: "A1" }
);
xlsx.writeFile(resultWorkbook, outputFile.replace("csv", "xlsx"));
