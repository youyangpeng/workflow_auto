
  function generateEmail() {

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Translators");

    const translator = sheet.getRange("F2").getValue();

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (data[i][0] == translator) {

        const source = data[i][1].trim();
        const target = data[i][2].trim();
        const email = data[i][3];
        const name = data[i][0].split("-")[0];
        
        let subject;
        let body;
        const sourceFr = languageMap[source];
        const targetFr = languageMap[target];
        
  
        const vowels = ["a", "à", "â", "e", "é", "è", "ê", "i", "î", "o", "ô", "u", "ù"];
        
        let sourceText;
        if (vowels.includes(sourceFr[0].toLowerCase())) {
        sourceText = `de l'${sourceFr}`;
        } else {
        sourceText = `du ${sourceFr}`;
        }

        let targetText;
        if (vowels.includes(targetFr[0].toLowerCase())) {
        targetText = `l'${targetFr}`;
        } else {
        targetText = `le ${targetFr}`;
        }

        if (source.includes("French") || target.includes("French")) {

          subject = "Demande de devis";

          body = `  
          Bonjour ${name},

          Pouvez-vous nous fournir un devis pour le document ci-joint ?

          Paire de langues : ${sourceText}>${targetText}

          Merci.
          `;

          } else {

            subject = `Quote Request - certified translation from ${source} into ${target}`;

            body = `
          
          Dear ${name},

          Could you please provide a quote for the attached document?

          Language Pair:  ${source}>${target}

          Thank you.
          `;

          }
          
        GmailApp.createDraft(
          email,
          subject,
          body
        );

        SpreadsheetApp.getUi().alert("Draft created!");

        return;
      }
    }

    SpreadsheetApp.getUi().alert("Translator not found.");

  }