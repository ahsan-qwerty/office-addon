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
    const data = await response.json();
    if (typeof data.fullText === "string" && data.fullText.length) {
      window.renderFullText(data.fullText);
      return;
    }
    const list = data.suggestions || [];
    window.renderSuggestions(list);
    await window.highlightSuggestions(list);
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
async function applyAllSuggestions(suggestions) {
  const sorted = [...(suggestions || [])].sort(
    (a, b) => (b.anchor?.length || 0) - (a.anchor?.length || 0)
  );
  for (const s of sorted) {
    // eslint-disable-next-line no-await-in-loop
    await applySuggestion(s.anchor, s.replacement);
  }
}

window.applyAllSuggestions = applyAllSuggestions;

async function replaceWholeDocument(newText) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertText(newText || "", Word.InsertLocation.replace);
    await context.sync();
  });
}

window.replaceWholeDocument = replaceWholeDocument;

const HIGHLIGHT_COLORS = ["#F0F8FF"];

async function highlightSuggestions(suggestions) {
  const list = Array.isArray(suggestions) ? suggestions : [];
  if (!list.length) return;
  await Word.run(async (context) => {
    const body = context.document.body;
    const resultsByIndex = [];
    list.forEach((s, idx) => {
      const results = body.search(s.anchor || "", {
        matchCase: false,
        matchWholeWord: true,
      });
      results.load("items");
      resultsByIndex.push({ idx, results });
    });
    await context.sync();
    resultsByIndex.forEach(({ idx, results }) => {
      const color = HIGHLIGHT_COLORS[0];
      results.items.forEach((r) => {
        r.font.highlightColor = color;
      });
    });
    await context.sync();
  });
}

async function clearHighlightsForAnchor(anchor) {
  if (!anchor) return;
  await Word.run(async (context) => {
    const body = context.document.body;
    const results = body.search(anchor, {
      matchCase: false,
      matchWholeWord: true, // make consistent with highlightSuggestions
    });
    results.load("items");
    await context.sync();

    results.items.forEach((r) => {
      r.font.highlightColor = null; // clears highlight
    });

    await context.sync();
  });
}

async function clearAllHighlights(suggestions) {
  const list = Array.isArray(suggestions) ? suggestions : [];
  for (const s of list) {
    // eslint-disable-next-line no-await-in-loop
    await clearHighlightsForAnchor(s.anchor);
  }
}

window.highlightSuggestions = highlightSuggestions;
window.clearHighlightsForAnchor = clearHighlightsForAnchor;
window.clearAllHighlights = clearAllHighlights;
