# Screenshot checklist

[Manual contents](README.md)

The first batch illustrates **13 of 28 topics: 9 complete, 4 partial, and 15 awaiting their first image**. Partial topics have a published overview plus a numbered placeholder for a lower-page companion image. No nonexistent image is embedded.

S05 is an annotated overview based on the supplied S07 screenshot; the unannotated original is retained. All other images in this batch are the supplied screenshots, unmodified. Click an image filename below to inspect the full-resolution asset.

## Capture guidelines

- Use the matching v0.9.9 testing build and current interface. Do not reuse an older screenshot just because the page has the same name.
- Prefer actual application screenshots over mockups. Keep the focus border where it helps explain navigation.
- Use consistent Windows scaling and window sizes. Capture full-screen examples at 1280×720 to demonstrate the intended small-screen layout.
- Make text readable on GitHub and in a later PDF. Use an overview plus a close crop when one image would make labels too small.
- Remove API keys, account identifiers, personal paths, private game/library names, and unrelated desktop content. Check masked key fields too before publication.
- Use short, consistent, non-private sample study content. Only include game artwork/screenshots you are entitled to publish.
- Keep the selected control and visible values consistent with the caption and nearby instructions.
- Save images under `docs/manual/images/`. Filenames below are reserved suggestions, not existing assets.
- Prefer PNG for interface text. Optimize image size without making labels blurry.

## Required shots

| ID | Status | Chapter / subject | Image filename | Coverage / next shot |
| --- | --- | --- | --- | --- |
| S01 | Complete | [Overall workflow](01-introduction.md) | [overall-workflow.png](images/overall-workflow.png) | Added: game, transcript/translation, and Explainer together. |
| S02 | Complete | [Extracted application](02-setup.md) | [extracted-application.png](images/extracted-application.png) | Added: extracted portable files with no personal parent path. |
| S03 | Complete | [API-key settings](02-setup.md) | [api-key-settings.png](images/api-key-settings.png) | Added: in-app key storage with both credentials masked. |
| S04 | Complete | [First capture and result](02-setup.md) | [first-capture-and-result.png](images/first-capture-and-result.png) | Added: original game dialogue and matching translated output; no region-selection outline. |
| S05 | Complete | [Desktop orientation](03-interface-and-controller.md) | [desktop-orientation.png](images/desktop-orientation.png) | Added: four numbered labels identifying sidebar, header Profile selector, page content, and footer. Original preserved as S07. |
| S06 | Partial | [Controls and bindings](03-interface-and-controller.md) | [controls-and-bindings.png](images/controls-and-bindings.png) | S06a added: action bindings. **Still needed S06b:** lower controller options, direct-action/D-pad switches, and detection status. |
| S07 | Partial | [Game Text page](04-game-text-translation.md) | [game-text-page.png](images/game-text-page.png) | S07a added: AI settings and capture actions. **Still needed S07b:** Formatting and Startup & advanced capture, including image-size limit. |
| S08 | Complete | [Capture choices](04-game-text-translation.md) | [capture-choices.png](images/capture-choices.png) | Added: desktop Capture region / Capture window menu. Its size limit lives in advanced capture settings, not in this popup. |
| S09 | Partial | [Audio setup and input check](05-audio-translation.md) | [audio-setup-and-input-check.png](images/audio-setup-and-input-check.png) | S09a added: AI/language and playback-device selection. **Still needed S09b:** completed input-check result and Start audio translation. |
| S10 | Pending | [Explanation settings](06-explanation.md) | `explanation-settings.png` | Show the AI choices, Explain latest text, and saving options. |
| S11 | Complete | [Explainer output](06-explanation.md) | [explainer-output.png](images/explainer-output.png) | Added: Japanese-source analysis beside the transcript and translation. |
| S12 | Complete | [Overlay appearance and typography](07-overlay-windows.md) | [overlay-appearance-and-typography.png](images/overlay-appearance-and-typography.png) | Added: Translator appearance, speaker color, font, size, and bold. |
| S13 | Complete | [Overlay positioning](07-overlay-windows.md) | [overlay-positioning.png](images/overlay-positioning.png) | Added: real positioning mode with move, resize, save, and cancel instructions. |
| S14 | Partial | [Terminology overview](08-terminology-overrides.md) | [terminology-overview.png](images/terminology-overview.png) | S14a added: enable switch and local-correction introduction. **Still needed S14b:** both glossary selectors and management controls. |
| S15 | Pending | [Glossary entry editor](08-terminology-overrides.md) | `glossary-entry-editor.png` | Show two or three fictional names and the source/replacement columns. |
| S16 | Pending | [Profile management](09-profiles.md) | `profile-management.png` | Show selection, Apply profile, Save current, New profile, and startup overlays. |
| S17 | Pending | [Per-game plugin setup](10-launchbox-and-big-box.md) | `per-game-plugin-setup.png` | Show both app icons, enable controls, Profile selections, readiness, and Save/Cancel. |
| S18 | Pending | [Expanded application locations](10-launchbox-and-big-box.md) | `expanded-application-locations.png` | Show browse controls and a useful ready/not-ready state; redact personal paths. |
| S19 | Pending | [Full-screen dashboard](10-launchbox-and-big-box.md) | `full-screen-dashboard.png` | Show Home, the help area, a focused tile, and navigation hints at 1280×720. |
| S20 | Pending | [Study Library overview](11-study-library.md) | `study-library-overview.png` | Show Library selection, search, table, selected source, version controls, and screenshot preview. |
| S21 | Pending | [Library dropdown and management](11-study-library.md) | `library-dropdown-and-management.png` | Show the separated Manage Study Libraries entry and the management options. |
| S22 | Pending | [Table filters and columns](11-study-library.md) | `table-filters-and-columns.png` | Show an example filtered result with Profile, Chapter, Japanese, Key grammar, and Versions. |
| S23 | Pending | [Metadata and current chapter](11-study-library.md) | `metadata-and-current-chapter.png` | Show future-entry chapter selection and a separate existing-entry details example. |
| S24 | Pending | [Study Reader](11-study-library.md) | `study-reader.png` | Show Japanese source, a learning section, version navigation, and screenshot context. |
| S25 | Pending | [Candidate review](11a-study-recommendations.md) | `candidate-review.png` | Show the Sentences/Vocabulary selectors, review scope, ratings/reasons, and explicit generation/add actions. |
| S26 | Pending | [Recommendation preferences](11a-study-recommendations.md) | `recommendation-preferences.png` | Show level/style, focus areas, and optional custom instructions. |
| S27 | Pending | [Anki connection and mapping](11b-anki.md) | `anki-connection-and-mapping.png` | Show a successful test, sample deck/note type, and the Japanese/Explanation field mapping. |
| S28 | Pending | [Card preview and explicit add](11b-anki.md) | `card-preview-and-explicit-add.png` | Show editable front/back, destination deck, optional screenshot, and Add to Anki. Use a small fictional study example. |

## Replace a placeholder

1. Capture and inspect the real application state.
2. Save the image with the corresponding proposed filename.
3. Replace that chapter's placeholder blockquote with an image and short caption.
4. Use meaningful alt text describing what the screenshot teaches.
5. Keep the ID in the caption/checklist so later updates can identify the shot.
6. Update the topic's Complete / Partial / Pending status and the counts above. Keep a topic Partial until its required companion shots are present.
7. Check the rendered chapter on GitHub and at the expected PDF reading size.

For example, once the actual image exists:

```markdown
![Game Text page with AI settings and the Capture & Translate action](images/game-text-page.png)

*Figure S07. Game Text settings and the capture actions.*
```

This code example intentionally does not render a missing image. Add the asset first, then use the same pattern in its chapter.

If one placeholder needs two images, append `-a` and `-b` to its filenames and use captions S07a/S07b, for example. Do not renumber the entire manual.
