# JRPG Translator v0.9.5

Version 0.9.5 turns the Study Library into a fuller post-game learning workflow.
The main addition is optional AnkiConnect integration for reviewing saved
explanations, selecting useful sentences or vocabulary, and creating cards with
their original gameplay context. The release also adds AI-assisted study tools,
spreadsheet export, richer metadata, and a broad reliability and interface pass.

## Highlights

### Anki-assisted review

The Study Library can now connect to a running Anki instance through
AnkiConnect. A read-only link check maps a unified JRPG Translator Profile to an
Anki deck or parent deck, note type, and fields, then finds exact normalized
Japanese matches across its subdecks.

From the Library or Reader you can:

- Add the complete explanation as an Anki card.
- Use Japanese without attached reading assistance on the front.
- Include a compact source screenshot on the back.
- Navigate long explanation cards with a generated section index and
  back-to-top links.
- Select a vocabulary line and create an editable vocabulary card using a
  cleaned dictionary-form front.
- Ask the selected AI provider to generate or regenerate a learner-friendly
  example sentence in Japanese, Japanese with readings, and the explanation
  language.

The new **Review for Anki** window collects sentence and vocabulary candidates,
supports new/backlog review, hides existing Anki matches, remembers finished
reviews, and can globally ignore vocabulary that should never be suggested
again.

### AI recommendations

Candidate recommendations are cached and can be generated or regenerated for a
chosen learner level and selection style. Optional settings cover vocabulary,
grammar, natural phrasing, reading comprehension, and free-form guidance.

Each candidate receives a score and short reason in the language of its saved
explanation. Vocabulary scoring considers the dictionary/base form even when
the saved line begins with an inflected form.

### Better Study Library organization

- Set an active chapter once and apply it automatically to new explanations;
  previously used chapters remain selectable and manageable.
- Properly formatted speaker headers can populate the Speaker field without
  guessing unmarked dialogue.
- New explanations store concise **Key grammar** metadata for quick review,
  filtering, tooltips, and export.
- Column order, width, and visibility are remembered.
- Right-click menus expose the same safe actions as the Library and candidate
  buttons.
- Export the Library and Anki review data to an Excel workbook with native date/
  time values.

### Reader improvements

Create a new explanation version from the Reader using a different provider,
model, or shared prompt without changing the main Explanation-tab model.
Navigation wraps between the first and last entries, retains the current section
when changing entries or versions, and avoids flashing the full explanation
during the switch.

### Translation and interface polish

- Speaker-name coloring now handles every formatted speaker header in multi-
  speaker output.
- Gemini/OpenAI failures are shown promptly in the overlays instead of leaving
  an indefinite hourglass.
- Dark-mode coverage now includes more dropdowns, dialogs, notifications,
  context menus, tables, scrollbars, and Anki/recommendation windows.
- Translucent tab switching no longer briefly exposes the desktop.
- Study windows, capture-region selection, tables, and high-DPI layouts received
  additional repaint, spacing, focus, and clipping fixes.

### Stability and platform updates

This release includes a project-wide AHK lifetime audit: timers and background
callbacks now verify that their GUI still exists, candidate refreshes tolerate
the window closing, controller polling is guarded against disconnects, and
temporary Study bridge output is isolated per operation.

The LaunchBox plugin now contains additional host-safety guards and validates
stored process IDs before cleanup. The bundled runtime target has moved to
64-bit Python 3.12 with refreshed active dependencies; unused legacy Google
Vision/client packages were removed and `openpyxl` was added for Excel export.

## Notes

- Anki integration requires the AnkiConnect add-on and a running Anki instance.
- Existing Study Libraries are upgraded in place; database backups are created
  before destructive Library operations.
- AI recommendations are optional and cached. Refreshing or reopening the review
  window does not repeat API calls unless recommendations are explicitly
  generated again.
- Existing explanations remain usable. Key grammar and automatically captured
  speaker/chapter metadata apply to newly generated entries.
