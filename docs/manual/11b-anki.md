# 11B. Anki connection, matching, and card creation

[← Recommendations](11a-study-recommendations.md) · [Contents](README.md) · [Next: Troubleshooting →](12-troubleshooting-and-advanced.md)

## Requirements

You need desktop **Anki**, a suitable deck and note type, and the **AnkiConnect** add-on. Anki must be running when JRPG Translator connects to it.

Install AnkiConnect using Anki's **Tools → Add-ons → Get Add-ons...** and code `2055492159`, then restart Anki. See the [AnkiConnect add-on page](https://ankiweb.net/shared/info/2055492159) for its current installation information.

JRPG Translator uses the local endpoint `http://127.0.0.1:8765`. This integration is with the running desktop application, not a direct login to AnkiWeb. Keep AnkiConnect local; there is no need to expose that port publicly.

Back up your Anki collection before experimenting with note types or a large import workflow.

## Configure the connection and fields

In the Study Library, choose **Anki... → Anki connection and link check...**.

1. Start Anki and choose **Test connection**.
2. Select the Study Profile/context to configure.
3. Choose the Anki deck or parent deck for link checking.
4. Choose the **Note type**.
5. Map its **Japanese field**.
6. Map its **Explanation field**.
7. Choose **Save mapping**.
8. Run **Refresh Anki status** when you want to compare Library entries with Anki.

> **Screenshot S27 — Anki connection and mapping (to be added).** Show a successful test, sample deck/note type, and the Japanese/Explanation field mapping.

Use the exact fields in your own note type. A note type can call them Front/Back, Japanese/Explanation, or something else; the mapping tells JRPG Translator which ones have the relevant meaning.

Ensure the note type's card template displays the mapped fields. Successfully adding a note to hidden or incorrectly mapped fields does not make a useful card.

The mapping is associated with the study/game context and can be remembered with its JRPG Translator Profile. The deck and note type themselves still live in Anki.

## Understand link checking

Link checking compares normalized Japanese source text against the configured deck/note-type/field scope. Selecting a parent deck includes its subdecks.

| Status | Meaning |
| --- | --- |
| Not checked | No current link-check result is available |
| Found in Anki | A matching source was found in the configured scope |
| Not found | The check did not find a matching source in that scope |

“Not found” is not proof that the material is absent from every deck in Anki. Check your deck, note type, Japanese field, and text differences first.

The manually editable **Added to Anki** metadata flag is separate from verified matching. Checking that box does not create an Anki note. Similarly, adding or changing notes directly in Anki may require refreshing status in the Library.

Link checking reads Anki data. Creating a note happens only through an explicit add action.

## Add an explanation or sentence

For an existing Library entry, select one explanation and choose **Anki... → Add selected explanation...**. You can also enter the card workflow from a selected sentence candidate.

1. Check the Japanese source/front and explanation/back.
2. Edit the proposed content to suit the card.
3. Choose the destination deck.
4. Decide whether to include a source screenshot when available.
5. Check any duplicate/match warning.
6. Choose **Add to Anki**.
7. Confirm success, then inspect the card in Anki.

> **Screenshot S28 — Card preview and explicit add (to be added).** Show editable front/back, destination deck, optional screenshot, and Add to Anki. Use a small fictional study example.

A good card usually asks one clear question. A complete multi-section explanation can be useful as reference on the back, but you do not have to keep every generated paragraph.

The app remembers sentence and vocabulary destination choices separately where configured. Always check the visible destination before adding, especially after switching Profiles.

## Add vocabulary

Open a vocabulary candidate and review its dictionary/base form, reading, meaning, and context. AI-generated explanations can extract an unusual form or choose the wrong meaning, so verify it against the sentence.

The vocabulary card workflow can generate an example sentence using the configured Explanation AI. **Generate example...** or regeneration is another AI request; adding a card does not require it if the existing material is sufficient.

If you generate an example, read both the Japanese and its explanation before accepting it. Generated examples are not quotations from the game and should not be treated as such.

## Duplicates and retries

The integration checks for matching material, but its matching rules cannot detect every conceptual duplicate. Two cards with different text can still teach the same word; two identical sources can intentionally belong to different contexts.

If an add operation times out or reports an uncertain result, inspect Anki and refresh matching before blindly retrying. The first request may have succeeded even if the confirmation was not received.

Do not delete notes or rewrite a deck merely to make the Library's status match expectations. Correct the mapping or refresh the check first.

## What this integration does not replace

Anki remains responsible for card templates, scheduling, reviewing, synchronization, and collection backups. JRPG Translator helps select and prepare material; it is not a replacement for managing your Anki collection.

The relevant local settings include `Settings\anki.ini`. Backing up that file preserves connection/mapping preferences, not the Anki collection itself.
