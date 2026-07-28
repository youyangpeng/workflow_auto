
  function quoterequestdraft() {


    const dashboard = SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Dashboard");

    const database =  SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Translators");

    const templates =  SpreadsheetApp.getActiveSpreadsheet()
                                .getSheetByName("Template");

    const translator = dashboard.getRange("A2").getValue();


    const data = database.getDataRange().getValues();

    const email_tpls = templates.getDataRange().getValues();

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

          subject = email_tpls[1][2].replaceAll("{{sourceText}}", sourceText)
          .replaceAll("{{targetText}}", targetText);

          body = [1][3].replaceAll("{{name}}", name)
          .replaceAll("{{sourceText}}", sourceText)
          .replaceAll("{{targetText}}", targetText);

          } else {

            subject = email_tpls[2][2].replaceAll("{{source}}", source)
          .replaceAll("{{target}}", target);;

            body =  email_tpls[2][3].replaceAll("{{name}}", name)
           .replaceAll("{{source}}", source)
          .replaceAll("{{target}}", target);

          }
        Logger.log (body)
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