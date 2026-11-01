# Blue Motion editable PowerPoint template

Visual reconstruction of the supplied motion presentation reference. Includes English and Persian editable PPTX/POTX files, an interactive standalone browser preview, an MP4 motion simulation, source configuration and a reproducible Python builder.

Start with `output/blue-motion.pptx` or `output/blue-motion-fa.pptx`. Use desktop Microsoft PowerPoint with Morph support and enable **Slide Show → Use Timings** for the intro. Click section titles or + in the outline to expand a section; click the same row again to collapse it. Slides 17–20 provide editable content pages.

Read **[the Persian customization guide](README-fa.md)** for the complete instructions, slide map, asset replacement and validation limits. See `reference/analysis.md` for observations, inferred techniques and unknowns.

```bash
python -m pip install -r requirements.txt
python tools/build_template.py
python tools/build_template.py --content content-fa.json --name blue-motion-fa
python tools/validate_pptx.py
```

Edit `content.json` / `content-fa.json` to update repeated content consistently. Optional `hero_path` and `logo_path` point to custom JPG/PNG assets relative to this folder. Rebuilding overwrites matching output files; preserve manual edits under a different filename first.

The MP4 and HTML are browser simulations, not recordings of native PowerPoint playback. Native Morph XML, package references and editable elements were checked; actual Microsoft PowerPoint playback was unavailable in this environment. The original PPTX and precise source assets cannot be recovered from the video. The blue artwork was generated for this reconstruction; the source creator's branding was replaced with editable sample branding.
