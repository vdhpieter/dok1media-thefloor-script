import fs from "fs";
import xlsx from "xlsx";

const questionsDir = "data/Questions";

const CategoryFolders = fs.readdirSync(questionsDir);

let result =
  "Vak,Recordnummer,Hoofdcategorie,Subcategorie,Vraag presentatie,Vraag graphics,Subcategorie presentatie,Volgorde,Antwoord graphics,Alle antwoorden,Antwoord dubbele laag,Vraag,Antwoord A,Antwoord B,Antwoord C,Meerkeuzeantwoord,Associatie 1,Associatie 2,Associatie 3,Associatie 4,Vraagsoort\n";

CategoryFolders.forEach((categoryFolder) => {
  if (categoryFolder === ".DS_Store") {
    return;
  }
  const questionFiles = fs.readdirSync(`${questionsDir}/${categoryFolder}`);

  questionFiles.forEach((questionFile) => {
    if (questionFile === ".DS_Store") {
      return;
    }
    console.log(questionFile);
  });
});
