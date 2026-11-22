# Bundled fonts (optional)

Drop OFL-licensed Hebrew TTF files here if you want the skill to embed a
specific brand font instead of relying on system fonts.

**Safe to bundle (SIL Open Font License):** Rubik, Heebo, Assistant,
Frank Ruhl Libre, David Libre, Noto Sans/Serif Hebrew, download from
[Google Fonts](https://fonts.google.com).

**Do NOT bundle** system fonts (David, Narkisim, Arial Hebrew), they are not
redistributable. Rely on them only when the target machine already has them.

You usually don't need anything here: macOS and Windows ship Hebrew fonts that
headless Chrome uses, and on Linux `sudo apt-get install fonts-noto-core`
(Debian/Ubuntu) covers it. See `../../references/hebrew-fonts.md`.
