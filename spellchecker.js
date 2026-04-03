async function loadRules() {
  const response = await fetch("grammar.xml");
  const xmlText = await response.text();

  const parser = new DOMParser();
  return parser.parseFromString(xmlText, "application/xml");
}

function getRules(xml) {
  const rules = [];
  const ruleNodes = xml.getElementsByTagName("rule");

  for (let rule of ruleNodes) {
    const find = rule.getElementsByTagName("find")[0].textContent;
    const replace = rule.getElementsByTagName("replace")[0].textContent;

    rules.push({ find, replace });
  }

  return rules;
}

async function applyRules() {
  const xml = await loadRules();
  const rules = getRules(xml);

  await Word.run(async (context) => {
    const body = context.document.body;

    for (let rule of rules) {
      const searchResults = body.search(rule.find, { matchCase: true });

      searchResults.load("items");
      await context.sync();

      searchResults.items.forEach(range => {
        // 1. Highlight the incorrect word
        range.font.highlightColor = "yellow";

        // 2. Suggest correction (from XML)
        range.insertComment("Suggestion: " + rule.replace);
      });
    }

    await context.sync();
  });
}