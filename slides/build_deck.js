// Build the project presentation:  node slides/build_deck.js   (run from the repo root)
// Numbers are read from slides/data.json (written by slides/slide_data.py from results/).
// Optional: NOTES=<path to a JSON of speaker notes keyed s01..s20> also writes a presenter copy with notes.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
// Optional: writes the theme colours into the .pptx (path to the pptx tooling's apply_theme.js).
const applyTheme = process.env.PPTX_SKILL ? require(process.env.PPTX_SKILL + "/scripts/apply_theme.js").applyTheme : null;

const D = JSON.parse(fs.readFileSync(path.join("slides", "data.json"), "utf8"));
const FIG = (f) => path.join("figures", f);
const NOTES = process.env.NOTES ? JSON.parse(fs.readFileSync(process.env.NOTES, "utf8")) : null;
const note = (s, key) => { if (NOTES && NOTES[key]) s.addNotes(NOTES[key]); };
const pct = (v) => (v === null || v === undefined ? "n/a" : v.toFixed(2) + "%");

const THEME = {
  name: "CT Surface Area",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2333", lt1: "FFFFFF", dk2: "14213D", lt2: "EEF1F6",
    accent1: "F0A202", accent2: "2A78D6", accent3: "EB6834", accent4: "1BAF7A",
    accent5: "EDA100", accent6: "6250D6", hlink: "2A78D6", folHlink: "6250D6",
  },
};
const HEX = THEME.colors;
// method colours, identical to the figures
const MC = { mc_grey: "2A78D6", cnn: "EB6834", pointnet: "1BAF7A", rf: "EDA100", mc_smooth: "E87BA4",
             mc_bin: "008300", lookup: "6250D6" };
const NAME = { face: "Voxel faces", mc_bin: "MC binary", mc_grey: "MC greyscale", mc_smooth: "MC smoothed",
               lookup: "MC + lookup", rf: "Random Forest", pointnet: "PointNet", cnn: "3D CNN" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Manas Jain";
pres.title = "Learned bias correction for surface-area estimation from CT-like voxel data";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333, H = 7.5, M = 0.6;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.9, w: W - 2 * M, h: 0.4, fontSize: 14,
                                bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 1.35, w: W - 2 * M, h: 2.2, fontSize: 38, bold: true,
                                color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.0, w: 6.4, h: 2.6, fontSize: 18,
                                color: "CADCFC", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.lt1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 8, h: 0.3, fontSize: 12, bold: true,
                                color: C.accent2, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.65, w: W - 2 * M, h: 0.85, fontSize: 30,
                                bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Manas Jain  |  PHC-351 Project 51", options: { x: M, y: H - 0.45, w: 6, h: 0.3, fontSize: 10,
                                color: "7A8496", margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.5, y: H - 0.45, w: 0.5, h: 0.3, fontSize: 10, color: "7A8496", align: "right" },
});

// ---------- helpers ----------
function content(sec, kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText(kicker.toUpperCase(), { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function card(s, x, y, w, h, name, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill || HEX.lt2 }, line: { color: fill || HEX.lt2 },
                                              rectRadius: 0.08, objectName: name });
}
function badge(s, x, y, label, name, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.46, h: 0.46, fill: { color: fill || HEX.dk2 }, line: { color: fill || HEX.dk2 },
                                 objectName: name });
  s.addText(String(label), { x, y, w: 0.46, h: 0.46, fontSize: 14, bold: true, color: C.background1, align: "center",
                             valign: "middle", margin: 0, isTextBox: true, objectName: name + " label" });
}
function rows(s, x, y, w, items, opts = {}) {
  // numbered rows: [{h: header, t: text}]
  const gap = opts.gap || 1.15;
  items.forEach((it, i) => {
    badge(s, x, y + i * gap, opts.labels ? opts.labels[i] : i + 1, `row ${i + 1} badge`, opts.fill);
    s.addText([{ text: it.h, options: { bold: true, color: C.text2, breakLine: true } },
               { text: it.t, options: { color: C.text1 } }],
              { x: x + 0.65, y: y + i * gap - 0.04, w: w - 0.65, h: gap - 0.1, fontSize: opts.fs || 15, valign: "top",
                margin: 0, isTextBox: true, objectName: `row ${i + 1} text`, paraSpaceAfter: 2 });
  });
}
function img(s, file, x, y, w, alt) {
  const sz = { "bias_vs_resolution.png": [1863, 775], "error_vs_noise.png": [1529, 717], "example_meshes.png": [1435, 574],
               "example_slices.png": [1440, 1000], "face_count_theory.png": [920, 600],
               "heldout_families.png": [1528, 800], "signed_error_box.png": [1551, 766], "training_curves.png": [920, 772] }[file];
  const h = w * sz[1] / sz[0];
  s.addImage({ path: FIG(file), x, y, w, h, altText: alt, objectName: file });
  return h;
}
function stat(s, x, y, w, big, small, color, name) {
  card(s, x, y, w, 1.45, name);
  s.addText(big, { x: x + 0.25, y: y + 0.12, w: w - 0.5, h: 0.75, fontSize: 36, bold: true, color: color || C.text2,
                   fontFace: THEME.headFontFace, margin: 0, isTextBox: true, objectName: name + " value" });
  s.addText(small, { x: x + 0.25, y: y + 0.85, w: w - 0.5, h: 0.5, fontSize: 13, color: C.text1, margin: 0, valign: "top",
                     isTextBox: true, objectName: name + " label" });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
                    legendFontFace: "+mn-lt", titleFontFace: "+mn-lt", catAxisLabelColor: "52514E",
                    valAxisLabelColor: "52514E", catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
                    dataLabelFontSize: 11, dataLabelColor: "1B2333", legendFontSize: 12, titleFontSize: 14,
                    titleColor: "1B2333" };
const MONO = "Consolas";
function callout(s, x, y, w, h, text, name, fill, color) {
  card(s, x, y, w, h, name, fill || "FDF3DC");
  s.addText(text, { x: x + 0.3, y, w: w - 0.6, h, fontSize: 15, color: color || C.text1, valign: "middle", margin: 0,
                    isTextBox: true, objectName: name + " text" });
}

// ---------- 1. title ----------
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
s.addText("PHC-351 COURSE PROJECT  |  PROJECT 51", { placeholder: "kicker" });
s.addText("AI-based 3D surface area estimation from CT-like voxel data", { placeholder: "title" });
s.addText([{ text: "Manas Jain", options: { bold: true, color: C.background1, breakLine: true } },
           { text: "Enrolment no. 24123021" }], { placeholder: "body" });
card(s, 7.35, 3.75, 5.4, 2.55, "mesh card", "FFFFFF");
img(s, "example_meshes.png", 7.45, 3.85, 5.2, "Reference surface of a rough blob next to binary and greyscale marching-cubes meshes");
note(s, "s01");

// ---------- 2. why it matters ----------
s = content("Introduction", "Why surface area matters", "Surface area is measured from CT scans in many fields");
[["Catalysts", "Reactions happen on the surface: more area, more activity."],
 ["Filters and rocks", "The pore surface controls how fluid flows through the material."],
 ["3D-printed lattices", "The area shows whether a printed part matches its design."]].forEach((c, i) => {
  const x = M + i * 4.1;
  card(s, x, 1.85, 3.85, 2.3, `use card ${i + 1}`);
  badge(s, x + 0.3, 2.1, i + 1, `use badge ${i + 1}`, MC.mc_grey);
  s.addText(c[0], { x: x + 0.95, y: 2.1, w: 2.75, h: 0.46, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0,
                    isTextBox: true, objectName: `use head ${i + 1}` });
  s.addText(c[1], { x: x + 0.3, y: 2.8, w: 3.3, h: 1.2, fontSize: 15, color: C.text1, valign: "top", margin: 0,
                    isTextBox: true, objectName: `use text ${i + 1}` });
});
card(s, M, 4.5, W - 2 * M, 2.15, "question card", HEX.dk2);
s.addText([{ text: "THE QUESTION", options: { fontSize: 13, bold: true, color: HEX.accent1, charSpacing: 2, breakLine: true } },
           { text: "How wrong is the surface area we measure from a CT scan, and can a learned model correct it?",
             options: { fontSize: 24, bold: true, color: "FFFFFF", fontFace: THEME.headFontFace } }],
          { x: M + 0.4, y: 4.6, w: W - 2 * M - 0.8, h: 1.95, valign: "middle", margin: 0, isTextBox: true, objectName: "question text" });
note(s, "s02");

// ---------- 3. what a CT scan gives us ----------
s = content("Introduction", "What a CT scan gives us", "Not a surface, but a grid of grey voxels");
img(s, "example_slices.png", M, 1.75, 5.7, "Central slices of one blob at 32 and 64 cubed and three noise levels");
rows(s, 6.75, 1.85, 6.0, [
  { h: "Voxels", t: "The scan is a 3D grid of small cubes. Each voxel stores one grey value." },
  { h: "Partial volume", t: "A voxel that is half inside the object comes out half grey." },
  { h: "Blur and noise", t: "The scanner blurs the image slightly and adds random noise (right column)." },
], { gap: 1.25 });
callout(s, 6.75, 5.65, 6.0, 1.05, "The true surface passes between voxel centres, so it has to be reconstructed. Every method for doing that adds some error.", "voxel callout");
note(s, "s03");

// ---------- 4. two classical methods ----------
s = content("Introduction", "Two standard methods", "Counting voxel faces, or fitting a triangle mesh");
const hMesh = img(s, "example_meshes.png", (W - 9.2) / 2, 1.65, 9.2, "Reference surface vs binary and greyscale marching-cubes meshes of a rough blob on a 32 cubed grid");
[["1  Count voxel faces", "Decide which voxels are inside, then count the faces on the outside of that block of cubes."],
 ["2  Marching cubes", "Build a triangle mesh through the grid and add up the triangle areas, on the black-and-white or the greyscale volume."]].forEach((c, i) => {
  const x = M + i * 6.15;
  card(s, x, 1.65 + hMesh + 0.2, 5.95, 1.3, `method card ${i + 1}`);
  s.addText([{ text: c[0], options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } },
             { text: c[1], options: { fontSize: 14, color: C.text1 } }],
            { x: x + 0.3, y: 1.65 + hMesh + 0.28, w: 5.4, h: 1.15, valign: "top", margin: 0, isTextBox: true, objectName: `method text ${i + 1}` });
});
note(s, "s04");

// ---------- 5. problem and objectives ----------
s = content("Introduction", "Problem and objectives", "Measure the error exactly, then learn to correct it");
card(s, M, 1.8, 4.6, 4.75, "problem card");
s.addText([{ text: "Problem", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
           { text: "Given a blurred, noisy voxel volume of one object, estimate the area of its true surface.", options: { breakLine: true } },
           { text: " ", options: { breakLine: true, fontSize: 8 } },
           { text: "Why it is hard", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
           { text: "Errors of 5% or more are common for small or noisy objects, often bigger than the differences people want to detect." }],
          { x: M + 0.3, y: 2.0, w: 4.0, h: 4.35, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "problem text" });
rows(s, 5.7, 1.85, 7.0, [
  { h: "Exact reference areas", t: "Shapes whose true area is known to better than 1 part in 10,000." },
  { h: "Measure the classical errors", t: "Voxel faces and three marching-cubes variants, across resolution and noise." },
  { h: "Learn a correction", t: "3D CNN, PointNet and Random Forest, compared fairly with a lookup table." },
  { h: "Test on unseen shapes", t: "Two shape families never used in training." },
  { h: "Practical advice", t: "When is a learned correction worth using?" },
], { gap: 0.95, fs: 14 });
note(s, "s05");

// ---------- 6. theory: voxel faces ----------
pres.addSection({ title: "Theory" });
s = content("Theory", "Theory 1: voxel faces", "Counting voxel faces is always about 50% too high");
const sph = Object.fromEntries(D.sphere.map((r) => [r.N, r]));
rows(s, M, 1.85, 6.2, [
  { h: "A staircase, not a surface", t: "A flat patch becomes steps. Step area = sum of its shadows on the 3 planes = |nx| + |ny| + |nz|." },
  { h: "Always too big", t: "Exactly 1 only when aligned with the grid, up to √3 = 1.73 on a diagonal, 1.5 on average: +50%." },
  { h: "Smaller voxels do not help", t: "The steps get finer, not flatter. Binary marching cubes has a smaller version of the same problem." },
], { gap: 1.3 });
stat(s, M, 5.85, 2.9, "+" + sph[128]["face_%"].toFixed(1) + "%", "voxel faces, clean sphere, 128³", C.text2, "stat faces");
stat(s, M + 3.2, 5.85, 2.9, "+" + sph[128]["mc_bin_%"].toFixed(1) + "%", "binary MC, clean sphere, 128³", C.text2, "stat mcbin");
img(s, "face_count_theory.png", 7.25, 1.9, 5.5, "Face-count area per unit true area against the polar angle of the surface normal");
note(s, "s06");

// ---------- 7. theory: curvature ----------
s = content("Theory", "Theory 2: curvature", "Curved surfaces come out slightly too small");
rows(s, M, 1.85, 6.0, [
  { h: "Flat surfaces: exact", t: "The grey values place the surface between voxels, so a flat surface comes out right." },
  { h: "Curved surfaces: too small", t: "Blur spreads the brightness outwards, so the half-brightness surface sits slightly inside the real one." },
  { h: "Predictable", t: "For a sphere: error ≈ −2σ²/R², with R the radius and σ² the total blur. Extra smoothing makes it about 3× bigger." },
], { gap: 1.4 });
const trow = (cells, hdr) => cells.map((c) => ({ text: String(c), options: { bold: !!hdr, color: hdr ? "FFFFFF" : "1B2333",
                                                  fill: { color: hdr ? HEX.dk2 : "FFFFFF" }, align: "center", fontSize: 14 } }));
const tbl = [trow(["Grid", "R (vox)", "Predicted", "Measured", "Smoothed pred.", "Smoothed meas."], true)];
D.sphere.forEach((r) => tbl.push(trow([`${r.N}³`, r.R_vox.toFixed(1), r["mc_grey_pred_%"].toFixed(2) + "%", r["mc_grey_%"].toFixed(2) + "%",
                                       r["mc_smooth_pred_%"].toFixed(2) + "%", r["mc_smooth_%"].toFixed(2) + "%"])));
s.addTable(tbl, { x: 6.85, y: 2.0, w: 5.9, colW: [0.8, 0.85, 1.05, 1.05, 1.1, 1.05], border: { type: "solid", pt: 0.5, color: "C9CED8" },
                  rowH: 0.45, fontFace: "Calibri", objectName: "curvature table" });
s.addText("Clean spheres: greyscale marching-cubes error predicted by the formula vs measured. The prediction is close wherever curvature dominates.",
          { x: 6.85, y: 4.45, w: 5.9, h: 0.8, fontSize: 12, italic: true, color: "52514E", margin: 0, isTextBox: true, objectName: "table caption" });
note(s, "s07");

// ---------- 8. theory: noise and the two biases ----------
s = content("Theory", "Theory 3: noise", "Noise adds area, so the two errors pull in opposite directions");
rows(s, M, 1.85, 6.0, [
  { h: "Fake specks", t: "Background voxels cross the threshold by chance. At the highest noise: about 112 specks on a 64³ grid." },
  { h: "Rough surfaces", t: "Noise roughens the real surface, and a rough surface always has more area." },
], { gap: 1.5 });
[["Curvature: area too small", "Strongest for small objects and coarse grids.", MC.mc_grey, "▼"],
 ["Noise: area too big", "Strongest for noisy scans and fine grids.", MC.cnn, "▲"]].forEach((b, i) => {
  const y = 1.85 + i * 1.55;
  card(s, 6.85, y, 5.9, 1.35, `bias card ${i + 1}`);
  s.addText(b[3], { x: 7.1, y, w: 0.6, h: 1.35, fontSize: 28, color: b[2], valign: "middle", align: "center", margin: 0,
                    isTextBox: true, objectName: `bias arrow ${i + 1}` });
  s.addText([{ text: b[0], options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
             { text: b[1], options: { fontSize: 14, color: C.text1 } }],
            { x: 7.85, y, w: 4.7, h: 1.35, valign: "middle", margin: 0, isTextBox: true, objectName: `bias text ${i + 1}` });
});
callout(s, M, 5.2, W - 2 * M, 1.3, "The size of each effect depends on the object and the scan, so no single correction factor can fix both. The correction has to be worked out object by object.", "bias callout", HEX.dk2, "FFFFFF");
note(s, "s08");

// ---------- 9. previous work ----------
pres.addSection({ title: "Previous work" });
s = content("Previous work", "Previous work", "Better fixed formulas, or learning without exact ground truth");
const cols = [
  { h: "Fixed estimators", t: ["Better weights for small voxel patterns: Mullikin and Verbeek 1993; Windreich et al. 2003; Lindblad 2005.",
                               "Triangle meshes beat voxel faces: Nyström et al. 2002."] },
  { h: "Convergence theory", t: ["Binary marching cubes does not converge to the true area, even for spheres: Kenmochi and Klette 2000.",
                                 "No local estimator converges on convex polytopes: Svane 2013."] },
  { h: "Learning on 3D data", t: ["3D CNNs on voxels (VoxNet, 2015) and PointNet on point sets (2017).",
                                  "DeePore (2020): CNN on micro-CT, but no exact ground truth for surface area."] },
];
cols.forEach((c, i) => {
  const x = M + i * 4.1;
  card(s, x, 1.85, 3.85, 3.6, `prev card ${i + 1}`);
  badge(s, x + 0.3, 2.05, i + 1, `prev badge ${i + 1}`);
  s.addText(c.h, { x: x + 0.9, y: 2.07, w: 2.8, h: 0.45, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true,
                   objectName: `prev head ${i + 1}` });
  s.addText(c.t.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c.t.length - 1 } })),
            { x: x + 0.3, y: 2.7, w: 3.3, h: 2.6, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 8,
              isTextBox: true, objectName: `prev text ${i + 1}` });
});
callout(s, M, 5.7, W - 2 * M, 1.05, "Gap: measure, against an exact reference, how much a learned model really reduces the error, and compare it with a simple baseline.", "gap card");
note(s, "s09");

// ---------- 10. method: data ----------
pres.addSection({ title: "Method" });
s = content("Method", "Method: data", "Synthetic shapes with a known true area");
rows(s, M, 1.85, 6.4, [
  { h: "Seven shape families", t: "Training: sphere, ellipsoid, cylinder, torus, smooth blob. Never trained on: superellipsoid (sharp edges), rough blob." },
  { h: "CT-like volumes", t: "Partial volume, blur and three noise levels, on 32³ and 64³ grids." },
  { h: "Split by shape", t: "No shape appears in both training and testing." },
  { h: "Exact reference area", t: "Exact formulas where they exist; a very fine triangle mesh for the blobs." },
], { gap: 1.2, fs: 14 });
[["2,400", "shapes"], ["14,400", "CT-like volumes (6 per shape)"], ["7", "shape families, 2 never trained on"], ["< 0.01%", "error of the reference area"]].forEach((t, i) => {
  stat(s, 7.3 + (i % 2) * 2.8, 1.85 + Math.floor(i / 2) * 1.75, 2.6, t[0], t[1], C.text2, `data stat ${i + 1}`);
});
callout(s, 7.3, 5.45, 5.4, 1.2, "Not marching cubes on a fine grid as the reference: it would put its own error into the ground truth.", "reference callout");
note(s, "s10");

// ---------- 11. method: models ----------
s = content("Method", "Method: the models", "The model learns a correction to marching cubes");
const flow = [["CT volume", "N³ grey voxels"], ["Greyscale MC", "area A_MC"], ["Model", "correction factor"], ["Corrected area", "A_MC × factor"]];
flow.forEach((f, i) => {
  const x = M + i * 3.1;
  card(s, x, 1.85, 2.6, 1.1, `flow ${i + 1}`, i === 2 ? "FDF3DC" : HEX.lt2);
  s.addText([{ text: f[0], options: { bold: true, color: C.text2, breakLine: true } }, { text: f[1], options: { color: C.text1 } }],
            { x: x + 0.15, y: 1.92, w: 2.3, h: 0.96, fontSize: 15, align: "center", valign: "middle", margin: 0, isTextBox: true,
              objectName: `flow text ${i + 1}` });
  if (i < 3) s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: x + 2.72, y: 2.25, w: 0.28, h: 0.3, rotate: 90,
                                                     fill: { color: "9AA5B8" }, line: { color: "9AA5B8" }, objectName: `flow arrow ${i + 1}` });
});
s.addText("The factor does not depend on size or orientation, so every training volume can be rotated and flipped into 48 versions.",
          { x: M, y: 3.05, w: W - 2 * M, h: 0.4, fontSize: 14, italic: true, color: "52514E", margin: 0, isTextBox: true, objectName: "target note" });
const models = [
  { h: "3D CNN", t: "Looks at the whole volume. 5 convolution blocks, 3.59M parameters.", c: MC.cnn },
  { h: "PointNet", t: "Looks at 2,048 points on the marching-cubes surface. 0.74M parameters.", c: MC.pointnet },
  { h: "Random Forest", t: "Uses 10 simple measurements: area ratios, number of pieces, noise estimate, shape.", c: MC.rf },
];
models.forEach((m, i) => {
  const x = M + i * 4.13;
  card(s, x, 3.65, 3.88, 2.0, `model card ${i + 1}`);
  s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 3.85, w: 0.32, h: 0.32, fill: { color: m.c }, line: { color: m.c }, objectName: `model dot ${i + 1}` });
  s.addText(m.h, { x: x + 0.7, y: 3.82, w: 3.0, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `model head ${i + 1}` });
  s.addText(m.t, { x: x + 0.25, y: 4.35, w: 3.4, h: 1.2, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `model text ${i + 1}` });
});
callout(s, M, 5.85, W - 2 * M, 0.9, "Baseline to beat: a lookup table with one fixed factor per resolution and noise level.", "baseline note");
note(s, "s11");

// ---------- 12. demo ----------
pres.addSection({ title: "Demo" });
s = content("Demo", "Live demo", "One model fixes errors in both directions");
card(s, M, 1.8, 7.7, 4.95, "terminal", "0F172A");
const tl = (t, c, b) => ({ text: t, options: { color: c || "D6DEEB", bold: !!b, breakLine: true } });
s.addText([
  tl("$ python src/demo.py --example blob --N 64 --noise 0.15", "F0A202", true),
  tl("synthetic blob, 64^3 grid, noise 0.15, diameter 55.7 voxels", "8FA3BF"),
  tl("Voxel-face counting          +83.72%"),
  tl("Marching cubes (binary)      +23.37%"),
  tl("Marching cubes (greyscale)    +9.25%"),
  tl("3D CNN correction             +0.24%", "7EE2B8", true),
  tl(" "),
  tl("$ python src/demo.py --example ellipsoid --N 32 --noise 0.05", "F0A202", true),
  tl("synthetic ellipsoid, 32^3 grid, noise 0.05, diameter 10.5 voxels", "8FA3BF"),
  tl("Marching cubes (greyscale)    -7.52%"),
  { text: "3D CNN correction             -0.96%", options: { color: "7EE2B8", bold: true } },
], { x: M + 0.3, y: 1.95, w: 7.2, h: 4.65, fontFace: MONO, fontSize: 13, valign: "top", margin: 0, paraSpaceAfter: 3,
     isTextBox: true, objectName: "terminal text" });
rows(s, 8.65, 1.95, 4.1, [
  { h: "Noisy blob", t: "Noise pushes greyscale MC up: +9.25%." },
  { h: "Small ellipsoid", t: "Curvature pulls it down: −7.52%." },
  { h: "Same CNN, both cases", t: "Corrected to +0.24% and −0.96%." },
], { gap: 1.45, fs: 15 });
note(s, "s12");

// ---------- 13. results: main ----------
pres.addSection({ title: "Results" });
const red32 = Math.floor(D.test.mc_grey["32"] / D.test.cnn["32"]), red64 = Math.floor(D.test.mc_grey["64"] / D.test.cnn["64"]);
s = content("Results", "Results", `Learned corrections cut the error ${red32} to ${red64} times`);
const order = ["mc_grey", "mc_smooth", "lookup", "pointnet", "rf", "cnn"];
s.addChart(pres.charts.BAR, [
  { name: "32³", labels: order.map((m) => NAME[m]), values: order.map((m) => D.test[m]["32"]) },
  { name: "64³", labels: order.map((m) => NAME[m]), values: order.map((m) => D.test[m]["64"]) },
], { x: M, y: 1.75, w: 7.6, h: 4.9, barDir: "bar", barGrouping: "clustered", chartColors: ["9AA5B8", HEX.dk2],
     showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", showLegend: true, legendPos: "b",
     showTitle: true, title: "Average error (%) on test shapes never seen in training", valGridLine: { color: "E4E7EC", size: 0.5 },
     catGridLine: { style: "none" }, valAxisMaxVal: 16, catAxisOrientation: "maxMin", barGapWidthPct: 60, ...chartText,
     objectName: "main results chart" });
stat(s, 8.55, 1.85, 4.2, pct(D.test.cnn["32"]), `3D CNN at 32³ (greyscale MC: ${pct(D.test.mc_grey["32"])})`, MC.cnn, "stat cnn");
stat(s, 8.55, 3.5, 4.2, pct(D.test.rf["64"]), `Random Forest at 64³ (greyscale MC: ${pct(D.test.mc_grey["64"])})`, "B07700", "stat rf");
s.addText([{ text: "The Random Forest matches the CNN. Its top inputs: number of pieces (counts noise specks) and noise estimate.", options: { breakLine: true } },
           { text: `PointNet is weaker (${pct(D.test.pointnet["32"])}, ${pct(D.test.pointnet["64"])}): its points come from a noisy mesh.` }],
          { x: 8.55, y: 5.15, w: 4.2, h: 1.6, fontSize: 13, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6,
            isTextBox: true, objectName: "results notes" });
note(s, "s13");

// ---------- 14. results: noise and size ----------
s = content("Results", "Results: noise and object size", "Robust to noise; the CNN is best on small objects");
const nl = ["σ = 0", "σ = 0.05", "σ = 0.15"], nm = ["mc_grey", "pointnet", "rf", "cnn"];
[32, 64].forEach((N, i) => {
  s.addChart(pres.charts.LINE, nm.map((m) => ({ name: NAME[m], labels: nl, values: [0.0, 0.05, 0.15].map((v) => D.noise[m][`${N}_${v === 0 ? "0.0" : v}`]) })),
    { x: M + i * 4.2, y: 1.75, w: 4.0, h: 4.0, chartColors: nm.map((m) => MC[m]), lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showTitle: true, title: `${N}³ grid, average error (%)`, showLegend: i === 1, legendPos: "b", valGridLine: { color: "E4E7EC", size: 0.5 },
      catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 10, ...chartText, objectName: `noise chart ${N}` });
});
stat(s, 9.2, 1.85, 3.55, pct(D.small.cnn), `3D CNN on objects under 16 voxels (greyscale MC ${pct(D.small.mc_grey)}, RF ${pct(D.small.rf)})`, MC.cnn, "stat small");
stat(s, 9.2, 3.55, 3.55, pct(D.noise.cnn["64_0.15"]), `3D CNN at 64³, highest noise (greyscale MC ${pct(D.noise.mc_grey["64_0.15"])})`, MC.cnn, "stat noise");
s.addText("The learned models barely notice noise. Object size matters more than grid size.",
          { x: M, y: 6.0, w: 8.2, h: 0.7, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: "noise note" });
note(s, "s14");

// ---------- 15. results: resolution ----------
s = content("Results", "Results: grid size", "The theory holds: errors vs grid size");
img(s, "bias_vs_resolution.png", (W - 10.6) / 2, 1.65, 10.6, "Mean signed error against grid size for binary, smoothed and greyscale marching cubes at three noise levels");
s.addText([{ text: "Binary MC gets worse with resolution under noise: +34% at 128³ (more voxels, more specks).", options: { bullet: true, breakLine: true } },
           { text: "Greyscale MC: too small at low resolution (curvature), too big at high noise.", options: { bullet: true, breakLine: true } },
           { text: "Learned models (squares) stay near zero at the two sizes they were trained on.", options: { bullet: true } }],
          { x: M, y: 6.15, w: W - 2 * M, h: 0.9, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: "sweep notes" });
note(s, "s15");

// ---------- 16. results: unseen shapes and failure ----------
s = content("Results", "Results: unseen shapes", "Works on new shapes; fails below two voxels");
img(s, "heldout_families.png", M, 1.75, 7.2, "Average error on the test split and the two held-out families at 32 and 64 cubed");
rows(s, 8.2, 1.85, 4.55, [
  { h: "Superellipsoids (never seen)", t: `CNN ${pct(D.super.cnn["32"])} / ${pct(D.super.cnn["64"])}, greyscale MC ${pct(D.super.mc_grey["32"])} / ${pct(D.super.mc_grey["64"])}.` },
  { h: "Rough blobs (never seen)", t: `Hardest case: CNN ${pct(D.rough.cnn["32"])} at 32³, too low. Roughness smaller than a voxel is not in the image.` },
  { h: "Failure case", t: "A torus with a 1.77-voxel tube: MC finds 42 of 142 units of area. No correction can fix that." },
], { gap: 1.55, fs: 14, labels: ["✓", "!", "✕"], fill: HEX.dk2 });
note(s, "s16");

// ---------- 17. which method ----------
pres.addSection({ title: "Conclusion" });
s = content("Conclusion", "Practical advice", "Which method should you use?");
const adv = [
  ["Any case", "Do not use voxel faces or binary MC for absolute area: large error that does not shrink.", "C0392B"],
  ["Large, clean objects", "Greyscale MC is already within about 1%.", MC.mc_grey],
  ["Noisy scans", "A learned correction (Random Forest or CNN) brings the error to about 1% or less.", MC.rf],
  ["Objects under 16 voxels", `The CNN is best: ${pct(D.small.cnn)}, against ${pct(D.small.mc_grey)} for greyscale MC.`, MC.cnn],
  ["Thinner than 2 voxels", "No method tested here is reliable.", "7A8496"],
];
adv.forEach((a, i) => {
  const y = 1.85 + i * 0.97;
  card(s, M, y, W - 2 * M, 0.82, `advice ${i + 1}`);
  s.addShape(pres.shapes.RECTANGLE, { x: M, y, w: 0.14, h: 0.82, fill: { color: a[2] }, line: { color: a[2] }, objectName: `advice bar ${i + 1}` });
  s.addText(a[0], { x: M + 0.4, y, w: 3.4, h: 0.82, fontSize: 16, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: `advice head ${i + 1}` });
  s.addText(a[1], { x: M + 3.9, y, w: W - 2 * M - 4.2, h: 0.82, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: `advice text ${i + 1}` });
});
note(s, "s17");

// ---------- 18. conclusion ----------
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("CONCLUSION", { placeholder: "kicker" });
s.addText("A learned correction removes most of the error, and theory explains where it comes from", { placeholder: "title" });
const concl = [
  `The classical errors are not random: they come from geometry, curvature and noise, and simple theory predicts them.`,
  `The best corrections reach ${pct(D.test.cnn["32"])} (32³) and ${pct(D.test.rf["64"])} (64³) on unseen test shapes, from ${pct(D.test.mc_grey["32"])} and ${pct(D.test.mc_grey["64"])}.`,
  `They still work on shape families never seen in training, with errors below 2%.`,
  `A simple Random Forest does as well as the 3D CNN in most cases.`,
];
s.addText(concl.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < concl.length - 1 } })),
          { x: M, y: 3.55, w: 12.1, h: 3.2, fontSize: 18, color: "E6ECF7", valign: "top", margin: 0, paraSpaceAfter: 10,
            isTextBox: true, objectName: "conclusion list" });
note(s, "s18");

// ---------- 19. limitations ----------
s = content("Conclusion", "Limitations and next steps", "What this study does not show yet");
const lims = [
  ["Simplified CT physics", "Noise added to the image, not to the X-ray projections; no artefacts such as beam hardening."],
  ["One object per volume", "Real materials have many touching structures; untested."],
  ["Fixed scan settings", "Two grid sizes, three noise levels, fixed blur."],
  ["Sub-voxel detail", "On rough surfaces the correction is a learned guess, not a measurement."],
  ["Simple shapes", "No thin sheets or shapes with many holes."],
  ["No real scan yet", "No tuning; the 64³ CNN trained only 15 epochs on a laptop."],
];
lims.forEach((l, i) => {
  const x = M + (i % 3) * 4.1, y = 1.85 + Math.floor(i / 3) * 2.05;
  card(s, x, y, 3.85, 1.85, `limitation ${i + 1}`);
  s.addText([{ text: l[0], options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
             { text: l[1], options: { fontSize: 14, color: C.text1 } }],
            { x: x + 0.3, y: y + 0.2, w: 3.25, h: 1.5, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true,
              objectName: `limitation text ${i + 1}` });
});
callout(s, M, 6.0, W - 2 * M, 0.75, "Next: a realistic CT simulation, volumes with many objects, and a real scan of an object with a known area.", "next steps");
note(s, "s19");

// ---------- 20. thanks ----------
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("PHC-351 PROJECT 51", { placeholder: "kicker" });
s.addText("Thank you", { placeholder: "title" });
s.addText([{ text: "Code, trained models, training and testing data and the article are included in the submission.", options: { breakLine: true } },
           { text: "Manas Jain  |  24123021" }], { placeholder: "body", w: 10 });
note(s, "s20");

(async () => {
  const out = process.env.OUT || path.join("slides", "Project51_ManasJain.pptx");
  await pres.writeFile({ fileName: out });
  if (applyTheme) await applyTheme(out, THEME);
  console.log("wrote", out);
})();
