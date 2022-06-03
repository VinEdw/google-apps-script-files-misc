function onOpen(e) {
  let ui = DocumentApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Format Selection as Chemical Equation', 'formatChemicalEquation')
    .addToUi();
}

function formatChemicalEquation() {
  const selection = DocumentApp.getActiveDocument().getSelection();

  function alignMatches(element, regex, textAlignment) {
    const text = element.getElement().editAsText();
    const textContent = text.getText().slice(element.getStartOffset(), element.getEndOffsetInclusive() + 1);
    let matches = [...textContent.matchAll(regex)];

    for (let j = 0; j < matches.length; j++) {
      let result = matches[j];
      let start = result.index + element.getStartOffset();
      let end = start + result[0].length - 1;
      text.setTextAlignment(start, end, textAlignment);
    }
  }

  if (selection) {
    const elements = selection.getRangeElements();
    for (let i = 0; i < elements.length; i++) {
      let element = elements[i];

      if (element.getElement().editAsText) {
        alignMatches(element, /\((s|l|g|aq)\)/g, DocumentApp.TextAlignment.SUBSCRIPT);
        alignMatches(element, /\d+[-+]/g, DocumentApp.TextAlignment.SUPERSCRIPT);
        alignMatches(element, /(?<=[\w]\s?)[-+]/g, DocumentApp.TextAlignment.SUPERSCRIPT);
        alignMatches(element, /(?<=[A-Za-z\)])\d+/g, DocumentApp.TextAlignment.SUBSCRIPT);
      }
    }
  }
}