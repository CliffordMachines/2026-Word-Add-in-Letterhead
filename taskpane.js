Word.run(async (context) => {

  const section = context.document.sections.getFirst();

  // HEADER
  const header = section.getHeader("Primary");
  const paragraph = header.insertParagraph(
    "Built to Perform, Designed to Last",
    Word.InsertLocation.end
  );

  paragraph.font.color = "#808080";
  paragraph.font.size = 11;
  paragraph.font.name = "Gotham Light";
  paragraph.alignment = Word.Alignment.centered;

  // PAGE MARGINS (points)
  section.topMargin = 106.30;
  section.bottomMargin = 913.68;
  section.leftMargin = 72.30;
  section.rightMargin = 72.30;

  // Ask Word what font it actually used
  paragraph.font.load("name");

  await context.sync();

  // FONT CHECK
  if (paragraph.font.name !== "Gotham Light") {
    Office.context.ui.displayDialogAsync(
      "data:text/html," + encodeURIComponent(`
        <html>
          <body style="font-family:Segoe UI; padding:16px;">
            <h3>Missing Font</h3>
            <p><strong>Gotham Light</strong> is not installed on this machine.</p>
            <p>Word used <strong>${paragraph.font.name}</strong> instead.</p>
            <p>Please install Gotham Light for correct letterhead branding.</p>
          </body>
        </html>
      `),
      { height: 30, width: 40 }
    );
  }

});