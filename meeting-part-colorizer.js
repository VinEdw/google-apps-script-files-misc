function changePartnerColors(letter = "V", color = "#6495ed") {
  function colorMatches(text, regex, color) {
    const textContent = text.getText();
    const matches = [...textContent.matchAll(regex)];
    console.log(matches);

    for (const match of matches) {
      const start = match.index;
      const end = start + match[0].length - 1;
      text.setForegroundColor(start, end, color);
    }

  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const bodyText = body.editAsText();

  colorMatches(bodyText, `${letter}:.+?\n`, color);
}