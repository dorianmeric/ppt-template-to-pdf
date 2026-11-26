# Assets

| File | Use |
| --- | --- |
| `logo.svg` | Mark for light backgrounds. |
| `logo-dark.svg` | Mark for dark backgrounds. GitHub uses it in dark mode through the `<picture>` element in the README. |
| `logo.png` | 512 × 512 PNG of the light mark. |
| `banner.html` | Source for `banner.png`. |
| `banner.png` | 1280 × 640 social preview image. |

The mark is a template slide with content bars, plus a PDF page in front of it. The page sits on a ring in the background colour, so the two shapes stay separate at small sizes.

Colours:

- Ink: `#1F2933` light, `#E6EDF3` dark.
- Accent: `#C00000` light, `#FF5C5C` dark.
- Dark background: `#0D1117`.

Regenerate the banner after editing `banner.html`:

```
msedge --headless --hide-scrollbars --window-size=1280,640 ^
       --screenshot=assets/banner.png assets/banner.html
```

Regenerate the PNG mark:

```
msedge --headless --hide-scrollbars --window-size=512,512 ^
       --default-background-color=00000000 ^
       --screenshot=assets/logo.png assets/logo.svg
```

Set the social preview: repository Settings → General → Social preview → Upload `assets/banner.png`. GitHub wants 1280 × 640 and under 1 MB.
