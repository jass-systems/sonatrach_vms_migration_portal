const ExcelJS = require('exceljs');

async function inspecterFormulaire(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  workbook.eachSheet((worksheet) => {
    console.log(`\n========================================`);
    console.log(` FEUILLE : ${worksheet.name}`);
    console.log(`========================================\n`);

    // Parcours de toutes les lignes et cellules renseignées
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        let contenu = cell.value;

        // Traitement du contenu selon le type (Formule, Texte enrichi, Objets)
        if (contenu && typeof contenu === 'object') {
          if (contenu.result !== undefined) contenu = contenu.result; // Résultat de formule
          else if (contenu.richText) contenu = contenu.richText.map(t => t.text).join('');
          else contenu = JSON.stringify(contenu);
        }

        console.log(
          `[Adresse: ${cell.address.padEnd(5)}] ` +
          `[Ligne: ${String(rowNumber).padStart(3)}, Col: ${String(colNumber).padStart(3)}] ` +
          `=> "${contenu}"`
        );
      });
    });

    // Identification des plages de cellules fusionnées (fréquentes dans les formulaires)
    if (worksheet.model && worksheet.model.merges && worksheet.model.merges.length > 0) {
      console.log('\n--- Zones / Champs fusionnés ---');
      worksheet.model.merges.forEach((range) => {
        console.log(`Zone fusionnée : ${range}`);
      });
    }
  });
}

inspecterFormulaire('formulaire_création_VM.xlsx').catch(console.error);