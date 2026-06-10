Office.onReady(() => {
  document.getElementById("addLetter").onclick = addHeaderText;
});

function addHeaderText() {
  Word.run(async (context) => {

    const sections = context.document.sections;
    context.load(sections);

    await context.sync();

    // Use first section (common case)
    const header = sections.items[0].getHeader("Primary");

    const paragraph = header.insertParagraph(
      "Built to Perform, Designed to Last",
      Word.InsertLocation.end
    );

    paragraph.font.color = "808080"; // light grey
    paragraph.font.size = 10;
    paragraph.font.name = "Calibri";
    paragraph.alignment = Word.Alignment.centered;

    await context.sync();
  });
}