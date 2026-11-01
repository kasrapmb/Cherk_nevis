# Video reconstruction evidence

## Source and limits

- Source: `VID_20261031_141515_145 (1).mp4`, 720 × 1280, 30 fps, duration 19.582 s.
- SHA-256: `c2bf1ba3dec87476a08ebc74ff99d811268d8794274a1279e09fb0947e7cf2f1`.
- Evidence is direct visual inspection of ffmpeg-extracted frames, not REA engine output. The requested REA skill's observations / inferences / unknowns discipline is used. No REA Evidence ID exists for these ordinary media-tool observations.
- The source is a handheld recording of a laptop, with perspective, reflections, warm white balance, and social overlays. The slide source, object tree, original assets, exact font, and animation settings are unavailable.
- All coordinates below refer to an estimated front-on slide, normalized to width/height 0–1; they are reconstruction guides rather than measured source geometry. A 16:9 slide is a reasonable inference from the displayed screen, not an exposed file property.

## Observed timeline

Times are seconds into the uploaded video, rounded to about 0.1–0.2 s. Fast motion between sampled frames remains approximate.

| Time | Directly observed behavior | Practical reconstruction |
|---|---|---|
| 0–3.5 | PowerPoint editing view, cover and additional slide thumbnails. | Exclude the laptop/editor and social captions from the template. |
| 4.0–4.4 | The screen darkens briefly entering presentation mode. | Not a required slide effect. |
| 4.5–5.3 | White field with a small central blue square containing a white geometric logo. | White opening state; center a replaceable logo tile. |
| 5.4–6.2 | A dark square behind the logo expands as a diamond, eventually covering the white field. Muted blue rectangular tiles extend horizontally around the logo. | Large rotated rectangle growth plus matched tile movement. Retain the centered logo. |
| 6.3–7.2 | Full dark navy field with blue mosaic tiles. Logo centered, brand and a small descriptor appear below. | Brand introduction state. |
| 7.3–8.6 | Logo shrinks and moves toward lower middle, brand moves beside it; a blue abstract image grows into a left panel. The remaining area is dark. | Matched object transformation from intro into cover layout. |
| 8.6–9.2 | Cover is assembled: large image left, title card right/top, brand and presenter below. | Editable cover state. |
| 9.3–10.1 | Cover content moves/fades away, large “Outline” shifts into the left part of the slide; a plain right panel takes over. | Matched title positioning and panel movement. |
| 10.3–11.2 | Right list appears in order: 01 Introduction, 02 Problem, 03 Solution, 04 Conclusion. Thin separators and white plus signs accompany rows. | Either staged list states or short entry effects. |
| 11.3–11.8 | Full outline holds: large title left; four rows right. | Outline overview state. |
| 11.9–12.7 | Title moves up/left and shrinks, rows move into the left portion. Abstract image appears in a tall right card with logo near its bottom-right. | Morph overview into compact accordion layout. |
| 12.8–13.6 | Row 01 expands to reveal three small lines of body text. Lower rows shift downward. White plus signs remain visible. | A matched slide state per expanded row; do not substitute a minus icon if aiming at the visible behavior. |
| 13.7–14.7 | Row 01 collapses, row 02 expands; right image changes from ribbon-like ridges to a rounded blue folded texture. The logo card tilts briefly during the image change. | Second accordion state and a separately replaceable image. |
| 14.8–15.2 | Second expanded state holds. | Optional hold. |
| 15.3–19.6 | Social end-card/Instagram overlay darkens the video. | Exclude from the editable presentation template. |

## Layout and styling observations

### Intro

- Logo tile is initially approximately centered `(0.50, 0.50)`, with side length around `0.10–0.15` of slide width depending on stage.
- Dark diamond grows from behind the tile, with a narrow light/gray edge visible in transition.
- Blue mosaic consists of low-contrast translucent-looking rectangles with darker/navy and brighter blue cells. It does not require a video background.

### Cover at about 9.0 s

- Dark navy full background, subtle blue variation; brighter blues occur in the image and one title line.
- Left image card: estimated `x=.055, y=.07, w=.43, h=.84`.
- Title card: estimated `x=.50, y=.09, w=.44, h=.53`, a slightly lighter navy than the base.
- Title begins around `x=.54, y=.22`: “Present with Clarity.” in white, “Design with Impact.” in blue; small “Pro Presentation” below.
- Bottom branding occupies approximately `x=.40–.72, y=.70–.89`; small square logo overlaps the bottom-right of the left image. Brand name and descriptor sit to its right.
- Presenter label is farther right, with a thin vertical divider separating it from the brand.
- White sans-serif text, titles bold, secondary labels smaller. Exact font and color values cannot be recovered from the recording.

### Outline overview at about 11.5 s

- Large “Outline” centered within left half at approximately `(0.25, .50)`.
- Left background is very dark navy; right half is a slightly lighter navy panel.
- Right list bounds roughly `x=.57–.94, y=.19–.80`.
- Row numbers blue; headings white. Plus signs right-aligned; faint thin horizontal rules under rows.
- Equal row spacing; no body paragraphs at this stage.

### Accordion at about 13.5 s

- Left text area extends from about `x=.07` to `.58`; image card occupies `x=.64–.96, y=.10–.91`.
- “Outline” is smaller and near `x=.07, y=.07`.
- Four rows start around `y=.25`, with expanded body under the active heading. Subsequent rows shift down to accommodate it.
- Expanded text is small, light gray, around three lines in the visible example. It is placeholder Latin text in the source.
- The geometric white logo remains on a dark square near the image's bottom-right, approximately `(0.92, .79)`.

## Inferences and implementation choices

- PowerPoint Morph is a strong practical match for continuous position/scale changes between intro, cover, overview, and accordion states. The recording does not prove the original used Morph.
- Shape animation or motion paths can also generate the diamond and tile reveal. A state-based Morph reconstruction keeps these objects editable and avoids baking them into a movie.
- Preserve the same object names (prefer `!!`-prefixed unique names), image identities where appropriate, and shape types across successive slides to encourage stable matching.
- Separate title, subtitle, number, heading, body, plus icon, separator, picture, and logo into editable objects. Duplicate the entire state slide before changing text positions; avoid grouping all content into a flattened bitmap.
- For each accordion state, keep the inactive rows present and adjust only text visibility/height and downstream row positions. Keep plus signs stable across all states.
- Use approximately 0.6–0.9 s Morph transitions for major layout changes and a shorter reveal for list build states. These durations are reconstruction choices, not recovered timing metadata.
- Provide a sequence that can be clicked through manually. If automatic timing is also provided, make the click-based version available for presentations.
- Replace the filmed blue artwork with a clean original/rebuilt abstract blue asset rather than including the laptop, reflections, or social overlay. Such an asset is visually analogous, not recovered original artwork.
- Changing a picture via PowerPoint's “Change Picture” retains useful object continuity better than deleting and inserting unrelated new shapes. Update repeated state slides consistently.
- Repeating the underlying editable logo geometry across states enables global recoloring; replacing the logo must be done consistently in all relevant states.

## Unknowns and validation boundary

- Unknown: original template file, source animation XML, authoring software version, exact easing, original image provenance, exact theme/master/font, precise click triggers and original total slide count.
- A visual reconstruction can match the demonstrated composition and motion logic. It cannot be claimed as the exact original template or byte-for-byte recovery.
- Static rendering verifies layout only. Actual Morph interpolation requires a compatible PowerPoint version and should be checked in its Slide Show view; alternative viewers may replace or ignore Morph.
