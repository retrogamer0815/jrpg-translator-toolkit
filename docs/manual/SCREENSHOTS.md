# Screenshot checklist

[Manual contents](README.md)

All 28 screenshots are **pending** in the first published manual. Each chapter contains a visible placeholder with a unique ID; no nonexistent image is embedded.

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

| ID | Chapter / subject | Proposed filename | What to show |
| --- | --- | --- | --- |
| S01 | [Overall workflow](01-introduction.md) | `overall-workflow.png` | Show a game with readable Translator and Explainer overlays. Use a short sample, with no personal account information. |
| S02 | [Extracted application](02-setup.md) | `extracted-application.png` | Show the executable alongside its supporting folders; crop the personal parent path. |
| S03 | [API-key settings](02-setup.md) | `api-key-settings.png` | Show the controls and status using masked or empty fields. Never use a real visible key. |
| S04 | [First capture and result](02-setup.md) | `first-capture-and-result.png` | Show the chosen game-dialogue region and the matching translated output. |
| S05 | [Desktop orientation](03-interface-and-controller.md) | `desktop-orientation.png` | Label the sidebar, header Profile selector, page content, and footer. |
| S06 | [Controls and bindings](03-interface-and-controller.md) | `controls-and-bindings.png` | Show detection status and examples of navigation/direct-action settings without implying that every controller has Xbox labels. |
| S07 | [Game Text page](04-game-text-translation.md) | `game-text-page.png` | Include AI settings, Capture & Translate, formatting, and the collapsed advanced section. Use a second crop if necessary at 720p. |
| S08 | [Capture choices](04-game-text-translation.md) | `capture-choices.png` | Show region, window, and maximum capture-image-size controls, with the help text for one focused option. |
| S09 | [Audio setup and input check](05-audio-translation.md) | `audio-setup-and-input-check.png` | Show device selection, Test audio, its result, and the start control. |
| S10 | [Explanation settings](06-explanation.md) | `explanation-settings.png` | Show the AI choices, Explain latest text, and saving options. |
| S11 | [Explainer output](06-explanation.md) | `explainer-output.png` | Show a short example with Japanese source and recognizable learning sections. |
| S12 | [Overlay appearance and typography](07-overlay-windows.md) | `overlay-appearance-and-typography.png` | Show color, opacity, font, size, and bold controls for one overlay. |
| S13 | [Overlay positioning](07-overlay-windows.md) | `overlay-positioning.png` | Show the positioning instructions and a window placed clear of dialogue and HUD elements. |
| S14 | [Terminology overview](08-terminology-overrides.md) | `terminology-overview.png` | Show the enable switch and the separate local-correction and Japanese-to-target-language sections. |
| S15 | [Glossary entry editor](08-terminology-overrides.md) | `glossary-entry-editor.png` | Show two or three fictional names and the source/replacement columns. |
| S16 | [Profile management](09-profiles.md) | `profile-management.png` | Show selection, Apply profile, Save current, New profile, and startup overlays. |
| S17 | [Per-game plugin setup](10-launchbox-and-big-box.md) | `per-game-plugin-setup.png` | Show both app icons, enable controls, Profile selections, readiness, and Save/Cancel. |
| S18 | [Expanded application locations](10-launchbox-and-big-box.md) | `expanded-application-locations.png` | Show browse controls and a useful ready/not-ready state; redact personal paths. |
| S19 | [Full-screen dashboard](10-launchbox-and-big-box.md) | `full-screen-dashboard.png` | Show Home, the help area, a focused tile, and navigation hints at 1280×720. |
| S20 | [Study Library overview](11-study-library.md) | `study-library-overview.png` | Show Library selection, search, table, selected source, version controls, and screenshot preview. |
| S21 | [Library dropdown and management](11-study-library.md) | `library-dropdown-and-management.png` | Show the separated Manage Study Libraries entry and the management options. |
| S22 | [Table filters and columns](11-study-library.md) | `table-filters-and-columns.png` | Show an example filtered result with Profile, Chapter, Japanese, Key grammar, and Versions. |
| S23 | [Metadata and current chapter](11-study-library.md) | `metadata-and-current-chapter.png` | Show future-entry chapter selection and a separate existing-entry details example. |
| S24 | [Study Reader](11-study-library.md) | `study-reader.png` | Show Japanese source, a learning section, version navigation, and screenshot context. |
| S25 | [Candidate review](11a-study-recommendations.md) | `candidate-review.png` | Show the Sentences/Vocabulary selectors, review scope, ratings/reasons, and explicit generation/add actions. |
| S26 | [Recommendation preferences](11a-study-recommendations.md) | `recommendation-preferences.png` | Show level/style, focus areas, and optional custom instructions. |
| S27 | [Anki connection and mapping](11b-anki.md) | `anki-connection-and-mapping.png` | Show a successful test, sample deck/note type, and the Japanese/Explanation field mapping. |
| S28 | [Card preview and explicit add](11b-anki.md) | `card-preview-and-explicit-add.png` | Show editable front/back, destination deck, optional screenshot, and Add to Anki. Use a small fictional study example. |

## Replace a placeholder

1. Capture and inspect the real application state.
2. Save the image with the corresponding proposed filename.
3. Replace that chapter's placeholder blockquote with an image and short caption.
4. Use meaningful alt text describing what the screenshot teaches.
5. Keep the ID in the caption/checklist so later updates can identify the shot.
6. Mark the entry as completed here and update the pending count.
7. Check the rendered chapter on GitHub and at the expected PDF reading size.

For example, once the actual image exists:

```markdown
![Game Text page with AI settings and the Capture & Translate action](images/game-text-page.png)

*Figure S07. Game Text settings and the capture actions.*
```

This code example intentionally does not render a missing image. Add the asset first, then use the same pattern in its chapter.

If one placeholder needs two images, append `-a` and `-b` to its filenames and use captions S07a/S07b, for example. Do not renumber the entire manual.
