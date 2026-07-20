# JRPG Translator v0.8.0

v0.8.0 makes model selection and overlay customization practical from a
controller while extending terminology consistency to screenshots,
explanations, and live audio.

## Highlights

- Screenshot Translation, Explanation, and Audio Translation now keep model
  lists appropriate to their individual workflows.
- The new `Add model` dialog can browse compatible OpenAI and Gemini models
  available to your API key or accept a model ID manually.
- Online model results are filtered by task, naturally sorted, cached locally,
  and fully navigable with arrow keys and Enter.
- Overlay transparency, font weight, and colors can now be configured without a
  mouse.
- The new controller color editor shows live hue, saturation, and brightness
  previews on informative gradient tracks.
- Translator and Explainer have independent Bold font settings, with improved
  Japanese pixel-font support for PixelMplus10 and PixelMplus12.
- Terminology Overrides now use case-insensitive matching with exact replacement
  casing and apply to screenshot translations and explanations. TL-to-TL rules
  also apply to live-audio output.
- Glossary and prompt editors now follow dark mode and open without selecting all
  existing text.
- Appearance and model controls have more consistent alignment, sizing,
  disabled states, and controller focus behavior.
- The advanced Paths tab is hidden by default, Debug mode defaults to off, and
  both remain available through `Settings/control.ini` for advanced users.

## Install Or Update

1. Download and extract the complete v0.8.0 portable ZIP.
2. Run `JRPG Translator.exe`.
3. Add API keys in the control panel or through Windows environment variables.

When updating an existing installation, keep a backup of your `Settings`
folder. The release contains maintained `default` prompt profiles, so preserve
any customized files that use the same names. Online model catalogs are cached
locally and can be refreshed from each `Add model` dialog.

See `CHANGELOG.md` for the complete list of changes.
