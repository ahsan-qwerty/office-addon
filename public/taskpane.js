/* global Office, Word */

async function improveSelection() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const response = await fetch("/api/improve", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ selectedText: selection.text }),
    });
    const { improvedText } = await response.json();

    selection.insertText(improvedText || "", Word.InsertLocation.replace);
    await context.sync();
  });
}

async function reviewWholeDoc() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const response = await fetch("/api/review", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: body.text }),
    });
    const { suggestions } = await response.json();

    window.renderSuggestions(suggestions || []);
  });
}

async function applySuggestion(anchor, replacement) {
  await Word.run(async (context) => {
    const body = context.document.body;
    const results = body.search(anchor, {
      matchCase: false,
      matchWholeWord: true,
    });
    results.load("items");
    await context.sync();
    if (results.items.length > 0) {
      results.items[0].insertText(replacement, Word.InsertLocation.replace);
      await context.sync();
    }
  });
}

// Expose to global for inline handlers
window.improveSelection = improveSelection;
window.reviewWholeDoc = reviewWholeDoc;
window.applySuggestion = applySuggestion;
