# Desktop orientation annotation (S05)

The unannotated source is [game-text-page.png](game-text-page.png), supplied by the project owner. The built-in image-generation/editing tool was used to add the four navigation callouts in [desktop-orientation.png](desktop-orientation.png). The annotated image is an instructional derivative; the original screenshot is retained separately and used for S07a.

Except for the S05 annotation and S03c privacy-redacted derivative described below, published images are the supplied screenshots copied without modification. S21's supplied filename duplicated the overview title; it is stored as `library-dropdown-and-management.png` to describe its actual content. S18 uses the corrected screenshot supplied on 20 September 2026, with personal account names redacted by the project owner in both user-specific paths.

The original S06 and S07 overview screenshots are retained as figures S06a and S07a. Their lower-page companions are S06b and S07b. The table records the corrected supplied filenames confirmed on 22 September 2026 and the companion screenshots added on 25 September 2026; the lower-page views do not replace the originals:

| Supplied filename | Published asset | Figure |
| --- | --- | --- |
| codex-clipboard-2551e760-ed10-42d9-8487-0d7f7845aa7a.png | [api-key-windows-shortcut.png](api-key-windows-shortcut.png) | S03b |
| codex-clipboard-c2bf0cf2-c12a-4d41-902a-ee69b589f5f8.png | [windows-user-api-variables.png](windows-user-api-variables.png) | S03c |
| S06 - controls-and-bindings.png | [controls-and-bindings.png](controls-and-bindings.png) | S06a |
| S06b - controls-and-bindings.png | [controller-options-and-detection.png](controller-options-and-detection.png) | S06b |
| S07 - game-text-page.png | [game-text-page.png](game-text-page.png) | S07a |
| S07b - game-text-page.png | [game-text-formatting-and-startup.png](game-text-formatting-and-startup.png) | S07b |
| S09b - audio-setup-and-input-check.png | [audio-input-check-result.png](audio-input-check-result.png) | S09b |
| S09c - audio-setup-and-input-check.png | [audio-start-stop-control.png](audio-start-stop-control.png) | S09c |
| S10 - explanation-settings.png | [explanation-settings.png](explanation-settings.png) | S10a |
| S10b - explanation-settings.png | [explanation-saving-and-startup.png](explanation-saving-and-startup.png) | S10b |
| S14b - terminology-overview.png | [terminology-selectors-and-local-corrections.png](terminology-selectors-and-local-corrections.png) | S14b |
| S21b - study-library-overview.png | [manage-study-libraries.png](manage-study-libraries.png) | S21b |
| S23b - metadata-and-current-chapter.png | [edit-explanation-details.png](edit-explanation-details.png) | S23b |
| S26b - recommendation-preferences.png | [regenerate-recommendations.png](regenerate-recommendations.png) | S26b |
| S26c - recommendation-preferences.png | [recommendation-customization.png](recommendation-customization.png) | S26c |

The S09–S26 companion screenshots added on 25 September 2026 are also unchanged originals. S21b now uses the supplied Manage Study Libraries replacement instead of the earlier Rename Library image; the redundant S21c placeholder was removed at the project owner's request. S26b shows the regeneration confirmation, and S26c shows the Customize dialog's study-focus areas and optional selection guidance. The unfilled S22b placeholder was also removed at the project owner's request; the existing S22a image remains. Superseded images remain recoverable from Git history.

The screenshot models, settings, Profile names, and controller bindings are examples, not a recommended or permanent default configuration. The full-screen dashboard image is the supplied 3840×2160 capture; it is not evidence of the 1280×720 layout.

## Windows API-key dialog privacy redaction (S03c)

The supplied Windows screenshot already covered credential values and personal path segments. The built-in image-editing tool was used to cover the remaining account name in the User variables heading. The published [windows-user-api-variables.png](windows-user-api-variables.png) is an edited derivative, not a pixel-identical original; the unredacted source was not added to the repository. The app's Windows shortcut screenshot (S03b) is unchanged.

### Redaction prompt

```text
Use case: precise-object-edit.
Asset type: redacted real Windows Environment Variables screenshot for a software user manual.
Input image: sole edit target, already supplied screenshot.
Primary request: add one solid opaque black rectangle covering only the account name immediately after 'User variables for' in the upper section heading. Fully hide all letters of that name. Leave the words 'User variables for' visible. Do not replace the name with other text.
Constraints: This is a privacy redaction, not a mockup. Keep the entire screenshot at its original size/aspect ratio and framing. Preserve all existing black redactions over API-key values and user paths exactly. Keep every other pixel, label, variable name (especially GEMINI_API_KEY and OPENAI_API_KEY), row, button, font, colour, border, layout and state unchanged. No additions except the single opaque black redaction over the account name. No text retyping, redesign, cropping, enhancement, or other edits.
```

## Desktop orientation annotation prompt (S05)

```text
Use case: precise-object-edit.
Asset type: annotated real software screenshot for a user manual, Figure S05.
Input image 1 is the sole edit target: the supplied actual JRPG Translator Game Text desktop screenshot.
Primary request: add exactly four small, polished instructional callout labels and short cyan leader arrows identifying these areas:
"1  Sidebar" -> the left navigation column containing Game Text, Audio, Explanation, Overlay windows, Study Library, Profiles, Settings.
"2  Header Profile selector" -> the existing top-right dropdown displaying "demo".
"3  Page content" -> the large central/right area containing AI settings and Capture & Translate.
"4  Footer" -> the bottom horizontal bar containing Translator: On, Explainer: On, Audio: Off, and Window options.
Keep the original screenshot, app logo, every existing word, number, font, glyph, control, color, panel, crop, and geometry unchanged. This is annotation, NOT redesign or a mockup. Do not re-typeset or redraw any original UI. Preserve the whole screenshot at its original aspect ratio and original resolution if possible. Do not blur or recolor the screenshot.
Use concise highly legible white sans-serif text in dark navy rectangular callout chips with thin cyan outlines, with discreet cyan leader lines/arrows. Add only those four labels, each once. Place labels in existing unused blank space without hiding UI labels or controls: Sidebar below the last sidebar item; Header Profile selector in empty space just below the header near the dropdown with arrow up to it; Page content in the unused area right of the Game Text Translation heading with arrow toward the main card; Footer in the wide empty bottom-bar gap between Audio: Off and Window options. Keep label edges comfortably inside the screenshot. No heading, watermark, legend, explanations, or other changes.
```
