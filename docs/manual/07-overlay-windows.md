# 7. Translator and Explainer overlay windows

[← Explanation](06-explanation.md) · [Contents](README.md) · [Next: Terminology Overrides →](08-terminology-overrides.md)

## Choose the right window

Open **Overlay windows**, then choose **Translator** or **Explainer**. Each has its own appearance and placement settings.

- The **Translator** displays screenshot translations and live audio output.
- The **Explainer** displays learning explanations.

These are not the desktop control panel. **Window options...** in the desktop footer changes the main interface, not the two output overlays.

> **Screenshot S12 — Overlay appearance and typography (to be added).** Show color, opacity, font, size, and bold controls for one overlay.

## Appearance

| Setting | Effect |
| --- | --- |
| Window color | Background behind the overlay text |
| Text color | Main text color |
| Speaker color | Translator speaker-name color when speaker highlighting is enabled |
| Opacity | Transparency of the entire overlay, including its text |
| Font | A font installed in Windows |
| Size (pt) / Font size | Text size |
| Bold | Heavier text rendering |

A low opacity can make text difficult to read even with a good background/text combination, because the letters become transparent too. Start with a high opacity, readable text, and enough contrast over both bright and dark game scenes.

Choose a font that contains the characters you need. A font suitable for Latin text may lack Japanese glyphs or use fallback glyphs with a different appearance. After moving the portable application to another PC, check fonts again; portable settings do not install fonts on that PC.

Changes to one overlay do not automatically synchronize the other.

## Move and resize

Use the **Move / Resize...** action for the intended overlay. The positioning mode lets you adjust the real window over the game.

In controller positioning mode, the left stick moves the window and the right stick resizes it. Confirm to keep the bounds or cancel to restore the previous position. Keyboard/mouse controls and any modifier used for keyboard resizing are shown in the positioning instructions.

> **Screenshot S13 — Overlay positioning (to be added).** Show the positioning instructions and a window placed clear of dialogue and HUD elements.

Test placement with both a short and a longer result. Leave enough width for the expected line length and enough height to avoid covering essential gameplay.

Save the result in the game's Profile if it should be restored next time. A Profile made on another resolution or monitor arrangement may need its bounds adjusted.

## Show, hide, and startup behavior

The footer and keyboard/controller actions provide quick overlay visibility controls. Hiding an overlay is useful while playing; it does not mean a running audio stream has stopped. Stop Audio Translation explicitly when you want to stop audio capture and AI usage.

Startup preferences determine which overlays should open when the application starts: none, Translator, Explainer, or both. Applying a Profile to an already-running application is not the same as restarting it and does not automatically run the startup-opening sequence.

The output overlays are designed to display results without acting like the main configuration window. Opening the desktop or full-screen control center can still take focus away from a game; some games pause when that happens.

## If an overlay disappears

Check:

1. Whether it is currently hidden.
2. Whether its opacity is too low or its text/background colors are too similar.
3. Whether saved bounds place it on a disconnected monitor.
4. Whether the game's exclusive-fullscreen mode covers normal desktop windows.
5. Whether an always-on-top/startup choice has only been set for the next opening.

For a game that covers desktop overlays, test borderless/windowed mode. If that solves it, the issue is the game's display mode rather than missing translation text.
