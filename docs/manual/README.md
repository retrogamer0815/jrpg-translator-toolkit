# JRPG Translator user manual

**For v0.9.9 (in testing).** Last checked against the application source: 18 September 2026.

JRPG Translator helps you translate Japanese game text and spoken dialogue while you play. Its optional learning tools explain the Japanese and let you build a searchable Study Library and Anki cards.

This manual describes the portable Windows application, its desktop and full-screen interfaces, and the optional LaunchBox / Big Box integration. Features, labels, and supported AI models can change during testing. The walkthroughs use the current interface rather than the older screenshots in the repository README.

## Start here

- **First installation:** read [Setup and your first translation](02-setup.md).
- **Playing with a controller:** continue with [Interface and controller navigation](03-interface-and-controller.md).
- **Setting up automatic game profiles:** see [Profiles](09-profiles.md), then [LaunchBox and Big Box](10-launchbox-and-big-box.md).
- **Learning Japanese:** read [Explanation](06-explanation.md) and [Study Library](11-study-library.md).
- **Something is not working:** use [Troubleshooting and advanced settings](12-troubleshooting-and-advanced.md).

You do not need to configure every feature before playing. One API key, a suitable image-translation model, a prompt, and a capture target are enough for a first text translation.

## Contents

1. [Introduction and key concepts](01-introduction.md)
2. [Download, setup, and your first translation](02-setup.md)
3. [The interface, keyboard shortcuts, and controllers](03-interface-and-controller.md)
4. [Game Text Translation](04-game-text-translation.md)
5. [Audio Translation](05-audio-translation.md)
6. [Explanation](06-explanation.md)
7. [Translator and Explainer overlay windows](07-overlay-windows.md)
8. [Terminology Overrides](08-terminology-overrides.md)
9. [Profiles](09-profiles.md)
10. [LaunchBox, Big Box, and the full-screen dashboard](10-launchbox-and-big-box.md)
11. [Study Library and the reader](11-study-library.md)
    - [Reviewing sentence and vocabulary recommendations](11a-study-recommendations.md)
    - [Anki connection, matching, and card creation](11b-anki.md)
12. [Troubleshooting, backup, and advanced configuration](12-troubleshooting-and-advanced.md)

## How to read this manual

A path such as `Settings\control.ini` is relative to the folder containing `JRPG Translator.exe`, unless stated otherwise. **Profile** means a saved set of JRPG Translator settings; a **Library** stores learning material. They are not interchangeable.

Menu paths such as **Settings → Controls** refer to the desktop interface. The full-screen dashboard exposes the same underlying settings through controller-friendly pages and pickers; differences are called out where useful.

Instructions that send content to an AI service are identified as AI operations. Changing a local setting, opening a saved explanation, or testing local audio input is not itself an AI request. A provider may charge for requests; consult its current billing information before use.

## Screenshots

The first two screenshot batches are included, with an annotated desktop overview and examples of Profiles, LaunchBox, Study Library, recommendations, and Anki. There are published images for 27 of the 28 topics; numbered callouts remain for missing companion views and Explanation settings. The [screenshot checklist](SCREENSHOTS.md) tracks completed, partial, and pending figures, filenames, and privacy checks.

## Maintaining this manual

These Markdown chapters are the primary documentation source. A PDF can be generated from the same chapters later, so the online and bundled manuals do not drift apart. No PDF is included with this initial manual.

When updating a release, check the version above, walkthroughs, default shortcuts, Profile scope, provider links, and screenshots together. AI availability and prices should be linked to official provider documentation rather than copied into a table that quickly becomes outdated.

[Next: Introduction →](01-introduction.md)
