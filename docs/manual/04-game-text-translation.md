# 4. Game Text Translation

[← Interface and controllers](03-interface-and-controller.md) · [Contents](README.md) · [Next: Audio Translation →](05-audio-translation.md)

## The basic loop

Choose a capture target, display the next line of dialogue, then use **Capture & Translate**. The application sends the captured image to the selected image-capable AI model and displays the result in the Translator overlay.

The game does not need to provide selectable text. Image quality still matters: very small characters, stylized fonts, animation, overlapping effects, or a region that cuts off text can reduce accuracy.

> **Screenshot S07 — Game Text page (to be added).** Include AI settings, Capture & Translate, formatting, and the collapsed advanced section. Use a second crop if necessary at 720p.

## Provider, model, and prompt

- **Provider** chooses the service used for screenshot translation.
- **Model** chooses the particular AI model for that service.
- **Prompt** controls the instructions and expected output style.

These selections are independent of the corresponding Explanation and Audio selections.

### Manage models

Use **Manage models...** to add or remove entries from the saved model list. The manager can obtain provider model choices and supports entering a model identifier when appropriate. Refresh an online list if it appears stale.

Use a model compatible with image input and the application's translation API. A text-only model, an audio-only model, an unavailable preview identifier, or a model excluded from your account can fail even if you add its name successfully.

Removing a model from the local list does not delete anything in your provider account.

### Manage prompts

Use the prompt's **Manage...** menu to edit the selected prompt, create a new one, or delete a prompt. Save an experimental prompt under a new name before changing a bundled one.

A translation-only prompt is useful for a compact overlay. A transcript/reading prompt is useful if you want to see the original Japanese or use the Explainer afterward.

**Prompt names currently matter.** The application recognizes `with_transcript` and `with_kanji_reading` in a translation prompt's name when selecting transcript-aware output processing. When creating a compatible custom prompt, preserve the appropriate naming convention and expected output structure. Merely asking for Japanese text inside an arbitrarily named prompt may not enable the intended processing.

Keep formatting instructions consistent with the bundled prompt you start from. Unstructured output can prevent speaker formatting or Japanese-source extraction from working as expected.

## Choose the capture target

Open **Change capture...** on the desktop, or **Capture...** in the full-screen interface.

### Select a region

A region captures fixed screen coordinates. It is useful when only the dialogue box is relevant.

Use the selection interface to position and size the region, then confirm. With controller selection, the left stick moves and the right stick resizes the selection; confirm/cancel follow the on-screen instructions. Mouse selection is also available.

If the game moves, changes resolution, or changes display arrangement, check the region again. A saved rectangle does not understand that a dialogue box has moved elsewhere.

### Select a window

A window target captures the selected game's window instead of a fixed desktop rectangle. Make sure you select the game, not the emulator's launcher, a settings window, or JRPG Translator itself.

Window titles and availability can change after a game restarts. Reselect the target if it no longer matches or the captured result is wrong.

**Selecting a region or window does not capture or translate anything by itself.** It only changes the target for subsequent actions.

> **Screenshot S08 — Capture choices (to be added).** Show region, window, and maximum capture-image-size controls, with the help text for one focused option.

## Three capture actions

| Action | What to use it for |
| --- | --- |
| Capture & Translate | Capture the current target and immediately request a translation |
| Make Capture | Save a capture and add it to the current in-memory batch, without making an AI request now |
| Translate Captures | Send all captures in that batch together in one AI request, without taking a fresh capture |

For multi-screen dialogue, use **Make Capture** while advancing through the relevant screens, then **Translate Captures**. A brief message reports how many shots are buffered. The batch is cleared when dispatched so repeated presses do not resend the old shots; an empty batch produces a “No buffered screenshots” message.

This is a current-session batch, not an import of every image in the screenshot folder. Finish the batch before restarting or switching to **Capture & Translate**: the single-capture action translates only its new image and clears the pending batch as part of dispatch. If a provider request later fails, do not assume the batch is still queued just because its image files remain on disk.

Captures can remain on disk. Translation should not be treated as a secure-delete operation. The startup cleanup setting described below is separate from making a request.

## Formatting options

### Highlight guessed subjects

Japanese often omits a subject that a target language may require. Compatible prompts can mark model-supplied subjects, and this option visually distinguishes those marked guesses.

This is a cue about the model's inference, not a confidence score or proof that every unmarked word is reliable. It depends on the prompt/output using the expected markers.

### Use speaker name color

This option highlights recognized speaker-name lines in the Translator overlay. Set the actual **Speaker color** under **Overlay windows → Translator**.

It does not assign a unique color to every character automatically. Speaker recognition and formatting depend on the captured content and prompt.

## Startup and advanced capture

The advanced section contains less frequently used capture/startup options. Configure these after the basic workflow works.

- **Capture image size limit (KB)** limits the image-file payload. It is not the overlay's font size or a pixel-width setting. A smaller limit can reduce detail; if characters become hard to read, use a tighter capture region or a more generous limit.
- **Startup overlay options** decide which output windows should open when the application starts. These do not replace current-session show/hide controls.
- **Clear screenshots on startup** controls cleanup of captures from previous sessions. Back up captures you want to keep before enabling cleanup.

Older documentation or configuration comments may refer to “PNG size.” PNG is the image-file format; the practical setting is the maximum size of the captured image sent for processing.

## Improve results without changing everything

If output is wrong, change one factor at a time:

1. Check that the actual captured image includes all dialogue.
2. Include a speaker name or a little surrounding context if relevant.
3. Capture after the game finishes drawing/animating the text.
4. Check the prompt's requested language and output format.
5. Try another compatible model if the image is readable but interpretation remains poor.
6. Add a [terminology rule](08-terminology-overrides.md) for a recurring name instead of manually repairing it every time.

More image area is not always more helpful. Capturing unrelated menus and text may make it harder for the model to identify the dialogue you want translated.
