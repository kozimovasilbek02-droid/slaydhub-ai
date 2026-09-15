---
name: slide-layout-intelligence
description: Master rules, geometric blueprints, layer hierarchies (Z-index), bounding box constraints (<=10% area limit), dynamic typography auto-scaling, and pure native vector PowerPoint reconstruction.
---

# Slide Layout Intelligence & Master Geometric Standards

Ushbu qo'llanma professional taqdimotlarni 100% toza vektor, aniq tipografiya va mukammal qatlamlar (Z-index) ierarxiyasi bilan yaratishning qat'iy standartlarini belgilaydi.

---

## 1. 6 Qatlamli Z-Index Ierarxiyasi va To'qnashuvni Oldini Olish (Anti-Collision)

Slaydga elementlarni joylashtirishda har doim quyidagi qat'iy Z-Index ketma-ketligiga rioya qilinishi SHART:

```
┌────────────────────────────────────────────────────────┐
│ Layer 5: Tahrirlanadigan Matn va Raqamlar (Topmost)    │  <-- Har doim eng ustida, hech qachon to'silmaydi
├────────────────────────────────────────────────────────┤
│ Layer 4: Shaffof PNG Cutout Rasmlar va Ikonkalar       │  <-- Fonidan tozalangan grafik aktivlar
├────────────────────────────────────────────────────────┤
│ Layer 3: Vektor Ko'rsatkichlar, Strelkalar va Tablar   │  <-- Bog'lovchi yo'llar, nishonlar, tablar
├────────────────────────────────────────────────────────┤
│ Layer 2: Asosiy Vektor Shakllar va Kartalar (Cards)    │  <-- MSO_SHAPE to'rtburchaklar, pufaklar, ovals
├────────────────────────────────────────────────────────┤
│ Layer 1: Yumshoq Soya (Ambient Drop Shadows)           │  <-- Shakl orqasidagi yumshoq soya qatlami
├────────────────────────────────────────────────────────┤
│ Layer 0: Slayd Asosiy Foni (Canvas Backdrop)           │  <-- Oq yoki mavzu foni
└────────────────────────────────────────────────────────┘
```

> [!IMPORTANT]
> **To'qnashuvga qarshi qoida (Anti-Collision Protocol):**
> Matn qutisi (`TextBox`) har doim eng oxirida qo'shiladi (Layer 5). Hech qanday shakl, karta, rasm yoki ikonka matn ustiga chiqib ketishi yoki uni to'sib qo'yishi qat'iyan man etiladi.

---

## 2. Matn Maydoni va 10% Cheklov Qoidasi (Text Bounding Box & <=10% Limit)

Matnlarni slaydda joylashtirishda quyidagi matematik talabga amal qilinadi:

1. **Maydon hisobi:** Asl namunadagi matn egallagan to'rtburchak maydon:
   $$Area_{orig} = Width_{orig} \times Height_{orig}$$
2. **Maksimal Ruxsat Etilgan Maydon:**
   $$Area_{max} = 1.10 \times Area_{orig} \quad (\text{Maksimal } +10\% \text{ gacha oshishi mumkin})$$
3. **Avtomatik Shriftsiz Kichraytirish (Dynamic Font Auto-Fitting):**
   - Agar matn kiritilganda u ajratilgan $Area_{max}$ dan oshsa yoki kartadan tashqariga chiqsa, shrift o'lchami (`Pt`) avtomatik ravishda mutanosib ravishda kichraytiriladi ($size = size \times 0.90$ yoki $size \times 0.85$).
   - Matn qatori noqulay uzilmasligi va shakl ichiga mukammal sig'ishi ta'minlanadi.
4. **Nol Margin Qoidasi:**
   `tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0` o'rnatilib, ortiqcha ichki bo'shliqlar yo'qotiladi.

---

## 3. 100% Toza PowerPoint Vektor Shakllarini Chizish Standarti

Har qanday infografika va geometrik shakl orqa fondan kesib olinmasdan, **to'g'ridan-to'g'ri PowerPoint mahalliy vektor shakllari (`MSO_SHAPE`)** yordamida yaratilishi SHART:

- **Lokatsiya Pinlari & Boshlar:** `MSO_SHAPE.OVAL` + `MSO_SHAPE.ISOSCELES_TRIANGLE` (180° rotation) + ichki oq doira.
- **Bosqichli Kartalar (Stepped Cards):** `MSO_SHAPE.ROUNDED_RECTANGLE` + o'ngga qaragan `MSO_SHAPE.ISOSCELES_TRIANGLE` (90° rotation) strelka teglari.
- **Pill Kartalar (Composite Pills):** Chapda `MSO_SHAPE.ROUNDED_RECTANGLE` (ikonka foni) + O'ngda rangli `MSO_SHAPE.ROUNDED_RECTANGLE` (sarlavha foni).
- **Aylanma Progress va Donutlar:** `MSO_SHAPE.OVAL` (qalin rangli stroke bilan) yoki `MSO_SHAPE.BLOCK_ARC`.
- **Streika va Bog'lovchilar:** `MSO_SHAPE.RIGHT_ARROW` yoki `MSO_SHAPE.RECTANGLE` chiziqlar.
- **Soyalar:** Asosiy shakl orqasiga `offset_x = 0.04"`, `offset_y = 0.04"`, rang `#E2E8F0` / `#00000015` bilan vektor soya chizish.

---

## 4. Rasmlar va Ikonkalarni Mukammal Fonidan Ajratish (Smart Cutout)

- **Fonni Aniqlash va Maskalash:** Fon piksellari rang masofasi ($d = \sqrt{\Delta R^2 + \Delta G^2 + \Delta B^2}$) bo'yicha aniqlanib, shaffof qilinadi.
- **Tashqi Artefaktlarni Tozalash:** Ikonka atrofidagi begona doira va chiziqlar avtomatik filtrlanib, faqat toza markaziy grafik qoldiriladi.
- **Monoxrom Vektor Ikonkalar:** Rangli yoki oq fon uchun ikonkalar kerakli rangga (`#2D3748` yoki `#FFFFFF`) moslashtiriladi.

---

## 5. Tipografiya Standartlari

| Element | Tavsiya etilgan Shrift | O'lcham (pt) | Qalinlik | Rang |
| :--- | :--- | :--- | :--- | :--- |
| **Asosiy Slayd Sarlavhasi** | Georgia / Montserrat | 28 - 34 pt | Bold | `#2D3748` / `#1A202C` |
| **Karta Sarlavhasi** | Montserrat / Segoe UI | 13 - 15 pt | Bold | Rangli fonda `#FFFFFF`, Oq fonda `#2D3748` |
| **Karta Tavsifi (Body)** | Calibri / Segoe UI | 9 - 10.5 pt | Regular | Rangli fonda `#FFFFFF`, Oq fonda `#64748B` |
| **Raqamli Natsenkalar (01..05)** | Montserrat / Arial | 14 - 18 pt | Bold | `#FFFFFF` |

---

### 6. Oppoq List Standarti va Vizual Tafsilotlar (Clean Canvas & Visual Details)
1. **Oppoq List Boshlang'ich Holati**:
   - Taqdimotdagi har bir slayd `prs.slide_layouts[6]` (toza bo'sh slayd) bilan noldan boshlanadi.
   - Slaydga hech qachon orqa foni bilan qirqilgan to'rtburchak raster rasmlar qo'yilmaydi.
2. **Sarlavhalar ustidagi brend aksenti (`add_header_accent`)**:
   - Har bir asosiy bo'lim sarlavhasining yuqorisiga gorizontal to'q sariq kapsula chiziqcha va nuqta (`— •`) qo'yiladi.
3. **Orqa fon suzuvchi simli to'rtburchaklar (`add_background_wireframe_squares`)**:
   - Slaydning bo'sh burchaklariga va struktura orqasiga nozik shaffof yumaloqlangan simli to'rtburchaklar (`MSO_SHAPE.ROUNDED_RECTANGLE`, `fill.background()`, `Pt(1.2)` chiziq) joylashtiriladi.
4. **Ustunlar orasidagi punktir chiziqlar (`add_dashed_vertical_divider`)**:
   - Metrika ustunlari orasiga punktir chiziqlar tortiladi.

### 7. Obyektlarni Toza Ajratish va Qatlamlar Tartibi (Pure Separation & Z-Index Assembly)
1. **Hech Qachon Matn Bilan Birga Qirqmaslik (Zero Ghost Text / Zero Debris)**:
   - Agar shakl orqa fondan qirqib olinsa, uning ustidagi yoki atrofidagi barcha matnlar (top text, bottom text, nishonlar) to'liq tozalanishi SHART.
   - Shakl ichidagi piksellardan tashqari barcha fonga oid otxodlar (faint wireframes, clipping debris) alfa 0 qilib tozalanadi.
2. **Obyektni To'liq Qirqish (No Clipping Bounds)**:
   - Ko'p arcli yoki radial shakllar (masalan, Slide 3 dial) qirqilganda tashqi chegaralar (o'ng tarafdagi chetki arklar) kesilib ketmasligi uchun yetarli kenglikda qirqiladi.
3. **6-Layer Clean Canvas Hierarchy (Z-Index)**
When constructing slides, always build from Layer 0 up to Layer 5 in order:
1. **Layer 0 (Blank Canvas)**: Pure `#FFFFFF` background.
2. **Layer 1 (Ambient Wireframes)**: Soft rounded squares (`MSO_SHAPE.ROUNDED_RECTANGLE`) with faint border `#E2E8F0` and no fill.
3. **Layer 2 (Core Visual Vector / Cutout)**: Unclipped pure vectors or transparent PNG cutouts (no ghost text, no edge cropping).
4. **Layer 3 (Structural Connectors)**: Crisp vertical stems (`MSO_SHAPE.RECTANGLE`) or curved arc lines passing precisely through node centers.
5. **Layer 4 (Floating Discs & Badges)**: Numbered discs (`MSO_SHAPE.OVAL`) with soft shadows or pure 360° circular badge cutouts.
6. **Layer 5 (Editable Typography)**: High-contrast native text boxes with strictly enforced $\le 10\%$ area bounds and dynamic font sizing (`add_constrained_textbox`).

## 8. Concentric Arc Trajectories (Yoyosimon Joylashuv)
- When numbered step indicators or cards accompany a circular dial, pie chart, or concentric ring graphic, **never force them into a straight vertical column**.
- Calculate their coordinates $(x_i, y_i)$ along the concentric circular arc matching the center $(c_x, c_y)$ and outer radius $R$ of the dial:
  $$x_i = c_x + R \cos(\theta_i), \quad y_i = c_y + R \sin(\theta_i)$$
- Connect the nodes with a matching curved circular arc line rather than a straight vertical line.
- Offset corresponding text labels to follow the curved $x_i$ displacement gracefully with zero overlapping.

## 9. Native Vector Rebuilding vs Raster Cropping
- For straightforward geometric shapes (waves, step paths, block arches, timelines, connected rings):
  - **Never use low-resolution or noisy raster crops** that suffer from pixelation or jagged edges.
  - Reconstruct them either as native PowerPoint vector shapes (`MSO_SHAPE.BLOCK_ARC`, `MSO_SHAPE.OVAL`, `MSO_SHAPE.RECTANGLE`) or render them via mathematical supersampled anti-aliasing (4x supersampling downscaled with Lanczos filtering) into crystal-clear, infinite-precision vector assets.

---

## 10. Radial Network Connectors vs Rotating Rectangles
- **Never use rotated rectangles** (`shape.rotation = angle`) to draw radial connector lines from a central hub to outer nodes. Rotating rectangles rotate around their center point, causing severe anchor detachment and misplaced line offsets.
- **Always use native straight connectors (`MSO_CONNECTOR.STRAIGHT`)**:
  ```python
  conn = slide.shapes.add_connector(
      MSO_CONNECTOR.STRAIGHT,
      Inches(cx_hub), Inches(cy_hub),
      Inches(x_target), Inches(y_target)
  )
  conn.line.color.rgb = col_rgb
  conn.line.width = Pt(3.5)
  ```
- **Z-Index Rule for Connectors**: Always add connectors to the slide *before* adding the center hub and outer node shapes, ensuring the lines emerge naturally from underneath nodes with zero visual overlap over icons or text.

---

## 11. Antialiased Segmented Color Wheels (Donut Charts)
- For multi-segment color wheels with $N$ distinct colored slices:
  - Generate an ultra-high resolution ($4\times$ supersampled) RGBA asset using mathematical arc slices ($\Delta\theta = 360^\circ / N - \text{gap}$), draw with `PIL.ImageDraw.pieslice`, punch out the inner circular hole with `fill=(0,0,0,0)`, and downscale with `LANCZOS` antialiasing.
  - Place on canvas at $(c_x - w/2, c_y - h/2)$, and overlay native vector circular disc (`MSO_SHAPE.OVAL`) with soft ambient drop shadow and title in Layer 5.

---

## 12. 4-Side Connected-Component Glyph Extraction (Zero Border Debris)
- When extracting iconography from templates, linear grid borders, underlines, or bounding boxes can contaminate the icon crop.
- Use connected-component labeling (`scipy.ndimage.label`) to identify and reject line artifacts (where aspect ratio $> 6$ or $w > 24$ with $h \le 4$, or components touching outer borders).
- Generate two standard variants:
  - `icon_w_{key}.png`: Pure white $(\text{RGB}=255,255,255)$ with alpha for dark/colored background shapes.
  - `icon_d_{key}.png`: Dark charcoal $(\text{RGB}=35,35,35)$ or themed colored glyphs for white circular badges.

---

## 13. Composite Multi-Part Cards (Tabs, Pills, and Number Badges)
- For composite cards (e.g. rounded body with left colored icon tab, right numbered badge `01..10`, or accent tabs):
  - **Layer 1**: Ambient drop shadow for the card (`MSO_SHAPE.ROUNDED_RECTANGLE`, fill `#E2E8F0` / `#CBD5E1`, line none).
  - **Layer 2**: Main white/light background rounded rectangle (`#FFFFFF` / `#F8FAFC`, border `#E2E8F0` `Pt(1)`).
  - **Layer 3**: Colored accent tab (`MSO_SHAPE.ROUNDED_RECTANGLE`) on the designated edge.
  - **Layer 4**: Circular icon badge (`MSO_SHAPE.OVAL`) + pure glyph picture (`icon_w` or `icon_d`).
  - **Layer 5**: Native text frame for title and description, plus standalone right number badge text frame (`f"{i+1:02d}"`) in bold white font.

---

## 14. Continuous Vertical Spine Alignment (Tree/Timeline Diagrams)
- For vertical lists with connected nodes, never draw individual disconnected stub lines between items.
- Draw one single continuous vertical line connector behind all nodes:
  ```python
  line = slide.shapes.add_connector(
      MSO_CONNECTOR.STRAIGHT,
      Inches(x_spine), Inches(y_first),
      Inches(x_spine), Inches(y_last)
  )
  line.line.color.rgb = hex_to_rgb('#CBD5E1')
  line.line.width = Pt(2.5)
  ```
- Place circular colored badge shapes (`MSO_SHAPE.OVAL`) directly over the spine at each row, cleanly masking the continuous line and creating a unified visual tree.

---

## 15. Multi-Layer Foreground Object Separation (Raster Photo vs Vector Background)
- **Problem Lesson (e.g., Food/Product/Thematic Slides)**: Placing a photo on a colored background as a single rectangular crop destroys slide quality and creates unsightly grey boxes.
- **Rule**:
  1. Extract the foreground photo/illustration as a transparent PNG cutout with crisp alpha channel masking (`alpha_cutout.png`).
  2. Rebuild the background shape natively in PowerPoint (`MSO_SHAPE.ROUNDED_RECTANGLE` or `MSO_SHAPE.OVAL`) using exact solid RGB fill and soft drop shadow.
  3. Place the transparent PNG photo on Layer 4 inside/over the vector shape, and add native editable text on Layer 5.

---

## 16. Vertical Banner & Side Tab Typography (Vertical Text Orientation)
- **Problem Lesson**: Side tabs and vertical ribbon banners often have labels oriented vertically (90° or 270°). Forcing them into horizontal single-character lines breaks typography.
- **Rule**:
  - Use `shape.rotation = 270` (or `90`) on dedicated text boxes or ribbon shapes.
  - Set `tf.word_wrap = False`, `tf.margin_top = tf.margin_bottom = 0`, and center alignment (`PP_ALIGN.CENTER`) so vertical banner text reads seamlessly from bottom to top or top to bottom.

---

## 17. Core Infographic vs Boilerplate Slide Classification
- **Problem Lesson**: SlideEgg and commercial template packs include 5-8 trailing auxiliary boilerplate slides (Icon sheets, Color Palette, "Thank You", "Contact Us", "About Us").
- **Rule**:
  - Automatically identify core infographic slides (Slides 1 to 5-7) that contain the primary data visualizations.
  - Dedicate full computational precision to the core slides.
  - Recognize icon sheets and auxiliary boilerplates so the system does not waste time attempting to force generic cards onto an icon directory slide.

---

## 18. UI Dashboard Non-Blocking & Responsive Standards (Streamlit / Desktop)
- **Streamlit Image Width Rule**: Never pass `width=None` to `st.image()`. Always use `width='stretch'`, `width='content'`, or integer pixel widths to maintain compatibility with modern Streamlit layout engines.
- **Lazy Drive Caching Rule**: Never execute synchronous O(N) recursive directory scans across large cloud/drive directories (11,000+ folders) during UI render loops. Cache folder trees in memory with `@st.cache_data` and execute live searches on-demand.

---

## 19. Dynamic Multi-Presentation Dispatching Architecture
- All precision builders must accept the target presentation directory dynamically via `sys.argv[1]` (with fallback to local source dir).
- Register builders in `slide_manager.py` (`builder_map`), and ensure preview PNG exports are namespaced per folder (`output/previews_{folder_name}/`) to eliminate any cross-presentation cache collision.

---

## 20. Tabular Calendar & Multi-Row Grid Matrices (Schedule & Planning Blueprints)
- **Geometry & Structure**:
  - Main container: `MSO_SHAPE.ROUNDED_RECTANGLE` with subtle 2pt theme border (`#00A896` / `#028090`) and solid white fill (`#FFFFFF`).
  - Day Header: 7 equal-width columns (`cw / 7.0`) using `MSO_SHAPE.RECTANGLE` with theme fill and bold white text (`Pt(12-14)`).
  - Date Matrix: 5-6 row grid with Sunday/Saturday column text highlighted in accent color (`#C0392B` / `#27AE60`) and weekdays in slate charcoal (`#1E293B`).
  - Highlighted Event Spans: Draw rounded accent cards (`#FDEBD0` / `#E8F8F5`) on Layer 2 *before* adding date text boxes on Layer 5 to indicate ongoing milestones or event windows.
  - Split Calendar Architecture: Left 35% dedicated to large month numerals (`Pt(70-90)`), thematic badge cards, and contact/notes sidebars; Right 65% reserved for the unobstructed calendar matrix.

---

## 21. Multi-Month Column Clusters & Pin-Note Matrices (Quarterly Blueprints)
- **Geometry & Multi-Column Partitioning**:
  - 3-Month / 4-Month Columns: Allocate symmetric 3.5" card widths with $0.65"$ horizontal gutters across a $13.333"$ canvas.
  - Circular Disc Card Architecture: Stack ambient shadow disc (`#E2E8F0`) $\to$ pure white disc (`#FFFFFF` with `#E2E8F0` 1pt stroke) $\to$ top colored accent tab (`MSO_SHAPE.ROUNDED_RECTANGLE`) $\to$ Day header row $\to$ 1-31 date matrix.
  - Realistic Sticky Note Cards: Layer 1 shadow offset ($x+0.08", y+0.08"$), Layer 2 tinted paper note body (`#FAF5FF` with `#E9D5FF` border), Layer 3 top centered circular pushpin badge (`MSO_SHAPE.OVAL`, diameter $0.4"$, dark theme fill `#5B21B6`), and Layer 5 sharp typography.
  - Angular Diagonal Header Banners: `MSO_SHAPE.RIGHT_TRIANGLE` placed at top-right ($x=4.5", y=0, w=8.833", h=3.8"$, `rotation=180`) with solid brand fill (`#991B1B` / `#0066CC`) to create modern corporate cover aesthetics.

---

## 22. 30-60-90 Day Strategy & Multi-Phase Velocity Architectures (Action Roadmaps)
- **Geometry & Phase Milestones**:
  - **Concentric Halo Discs**: Layered 3.5pt outer halo ring (`#F8FAFC`) + 1.5pt inner white hub disc (`Pt(22)` bold font) + matching colored pill tabs (`Pt(12)` bold white text).
  - **Overlapping Disc Top Badges**: Draw outlined container card (`MSO_SHAPE.ROUNDED_RECTANGLE`, 2.5pt stroke) on Layer 2, then place circular phase badge (`MSO_SHAPE.OVAL`, $1.3" \times 1.3"$) overlapping the top border with a 3pt solid white outline on Layer 4 to create a 3D badge depth effect.
  - **Horizontal Capsule Ribbons**: Left 360° circular milestone disc ($1.5" \times 1.5"$) integrated into a rounded horizontal ribbon body ($w=9.2", h=1.3"$) with 2-tier white header and tinted description typography.
  - **Slanted Parallelogram Blocks & Process Chevrons**: Use `MSO_SHAPE.PARALLELOGRAM` and `MSO_SHAPE.CHEVRON` in sequential high-contrast triads (Amber `#F59E0B` $\to$ Rose `#E11D48` $\to$ Teal `#0D9488`) with bottom aligned legend descriptions.

---

## 23. Neural Alpha Foreground Isolation & Vector Geometry Compositing (`rembg` + `shapely`)
- **Neural Object Cutout Pipeline**:
  - Use `rembg.remove()` for isolating photographs, product cutouts, food items, and complex graphics into transparent PNG alphas, eliminating background color clashes.
  - Apply `scipy.ndimage` connected-component thresholding on monochrome glyph icons to strip non-icon debris and bounding frame lines before placement.
  - Utilize `shapely.geometry` for 2D computational geometry, polygon intersections, and bounding box validation to guarantee zero text container overflow ($\le 1.10 \times \text{original area}$).

---

## 24. Executive Pillar Sidebars, Continuous Progress Beams & Neo-Brutalist 3D Offset Cards
- **Geometry & Hierarchy**:
  - **Left Pillar Sidebar**: Allocate $2.0"$ off-white pillar card (`#F8FAFC`) with stacked vertical titles (`Georgia Pt(26)`), accented by a right vertical brand strip (`#6366F1`, $w=0.12"$).
  - **Continuous Progression Beam**: Draw a $0.15"$ thick horizontal connecting bar spanning across all phase columns with circular node dots ($0.24" \times 0.24"$, `#FFFFFF` fill with `#6366F1` 1.5pt border) centered over each column axis.
  - **Bottom Dual-Halo KPI Badges**: Overlap bottom border with circular percentage badges ($1.2" \times 1.2"$, 3pt white outline, `Pt(15)` bold white text) to emphasize metric targets (e.g. 50%, 30%, 48%).
  - **Neo-Brutalist 3D Chromatic Offset Cards**: Layer 1 solid saturated offset shadow base ($x-0.12", y+0.12"$, `#FFC107` / `#F43F5E` / `#06B6D4`), Layer 2 pastel foreground rounded card (`#FEF9C3` / `#FFE4E6` / `#CFFAFE`), and Layer 3 top milestone node sphere ($1.5" \times 1.5"$) with large bold numerals (`Pt(28)`).

---

## 25. Classical Architectural Metaphors (Pillars, Pediments & Stepped Podiums)
- **Geometry & Classical Metaphor Construction**:
  - **Pediment (Roof Triangle)**: `MSO_SHAPE.ISOSCELES_TRIANGLE` ($w=10.333", h=1.8"$) with subtle slate-grey fill (`#E2E8F0`) and center circular architectural medallion (`MSO_SHAPE.OVAL`, $1.3" \times 1.3"$).
  - **Stepped Foundation Podium**: 2-tier stacked rectangles at bottom ($h=0.25"$ each, bottom base $w=9.733"$, upper step $w=9.133"$) to provide visual ground anchor.
  - **Tripartite Structural Columns**:
    - Capital (Top block): `MSO_SHAPE.RECTANGLE` ($w=2.8", h=0.35"$, theme color).
    - Shaft (Column body): `MSO_SHAPE.ROUNDED_RECTANGLE` ($w=2.4", h=2.25"$, theme color) with center circular white icon hub ($0.8" \times 0.8"$) and crisp white typography.
    - Plinth (Bottom base): `MSO_SHAPE.RECTANGLE` ($w=2.8", h=0.35"$, theme color).

---

## 26. Rotated Diamond Quadrant Pinwheels (4P / 4C Strategy Matrices)
- **Geometry & Radial Diamond Array**:
  - 4 outer rounded rectangles rotated 45° (`MSO_SHAPE.ROUNDED_RECTANGLE`, $2.0" \times 2.0"$, `rotation=45`) arrayed radially around center $(x_c, y_c)$ at $(\pm 1.6", 0)$ and $(0, \pm 1.6")$.
  - 4 inner colored core diamond tabs (`rotation=45`, $1.0" \times 1.0"$) in high-contrast marketing triads (`#00A896` Emerald, `#E65100` Orange, `#D84315` Terracotta, `#1565C0` Blue).
  - Clean split canvas: Left 35% reserved for bold 2-line title (`Georgia Pt(36)`), Right 65% dedicated to the unobstructed diamond pinwheel.
