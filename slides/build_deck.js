// Build the project presentation:  node slides/build_deck.js   (run from the repo root)
// Numbers are read from slides/data.json (written by slides/slide_data.py from results/).
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
// Optional: writes the theme colours into the .pptx (path to the pptx tooling's apply_theme.js).
const applyTheme = process.env.PPTX_SKILL ? require(process.env.PPTX_SKILL + "/scripts/apply_theme.js").applyTheme : null;

const D = JSON.parse(fs.readFileSync(path.join("slides", "data.json"), "utf8"));
const FIG = (f) => path.join("figures", f);
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

// ---------- 1. title ----------
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
s.addText("PHC-351 COURSE PROJECT  |  PROJECT 51", { placeholder: "kicker" });
s.addText("Learned bias correction for surface-area estimation from CT-like voxel data", { placeholder: "title" });
s.addText([{ text: "Manas Jain", options: { bold: true, color: C.background1, breakLine: true } },
           { text: "Enrolment no. 24123021", options: { breakLine: true } },
           { text: "Department of Physics, IIT Roorkee" }], { placeholder: "body" });
card(s, 7.35, 3.75, 5.4, 2.55, "mesh card", "FFFFFF");
img(s, "example_meshes.png", 7.45, 3.85, 5.2, "Reference surface of a rough blob next to binary and greyscale marching-cubes meshes");
s.addNotes("Hello, I am Manas Jain, enrolment number 24123021, and this is my PHC-351 project, number 51: AI-based 3D surface area estimation. The question is simple: when we measure the surface area of an object from a CT scan, how wrong is the number, and can a learned model correct it? The picture on the right already shows the problem: the true surface, a staircase mesh from a binary volume, and a smooth mesh from the greyscale volume.");

// ---------- 2. introduction ----------
s = content("Introduction", "Introduction", "CT surface area is biased, and the bias depends on the object");
rows(s, M, 1.85, 5.9, [
  { h: "Why it matters", t: "Surface area drives catalyst activity, flow in porous media and the quality of 3D-printed lattices." },
  { h: "How it is measured", t: "Count the exposed faces of segmented voxels, or extract a triangle mesh with marching cubes and sum its area." },
  { h: "What goes wrong", t: "Both are biased. The error depends on resolution, noise, curvature and thin features at the same time." },
  { h: "This project", t: "An exact synthetic benchmark, and learned models that correct the marching-cubes area per object." },
], { gap: 1.2 });
const hMesh = img(s, "example_meshes.png", 6.85, 2.1, 5.9, "Reference surface vs binary and greyscale marching-cubes meshes of a rough blob on a 32 cubed grid");
s.addText("Same rough blob on a 32³ grid: the binary mesh is a staircase (+6.6%), the greyscale mesh is smooth but loses fine detail (-2.0%).",
          { x: 6.85, y: 2.1 + hMesh + 0.1, w: 5.9, h: 0.7, fontSize: 12, italic: true, color: "52514E", margin: 0, isTextBox: true,
            objectName: "mesh caption" });
s.addNotes("Surface area matters for catalysts, filters and additively manufactured lattices. From a CT volume it is measured in one of two ways: counting voxel faces, or running marching cubes and adding up triangle areas. Both are biased, and the size and even the sign of the error depend on resolution, noise, curvature and thin features. So a single correction factor cannot fix it. My project builds a benchmark where the true area is known exactly, and trains models that correct the marching-cubes area object by object.");

// ---------- 3. problem and objectives ----------
s = content("Introduction", "Problem and objectives", "Measure the bias exactly, then learn to correct it");
card(s, M, 1.8, 4.6, 4.6, "problem card");
s.addText([{ text: "Problem", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
           { text: "Given a CT-like greyscale volume of one object, estimate the area of its true continuous surface.", options: { breakLine: true } },
           { text: " ", options: { breakLine: true, fontSize: 8 } },
           { text: "Why it matters", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
           { text: "Errors of 5% or more are common for small or noisy objects. That is larger than many of the differences surface-area measurements are used to detect." }],
          { x: M + 0.3, y: 2.0, w: 4.0, h: 4.2, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "problem text" });
rows(s, 5.7, 1.85, 7.0, [
  { h: "Exact reference areas", t: "Seven shape families, reference area accurate to better than 1 part in 10,000." },
  { h: "Quantify classical bias", t: "Voxel faces and three marching-cubes variants across resolution and noise." },
  { h: "Learned corrections", t: "3D CNN and PointNet vs a lookup table and a Random Forest, at matched resolution and noise." },
  { h: "Generalisation", t: "Test on shape families never seen in training." },
  { h: "Recommendations", t: "When is a learned correction worth using?" },
], { gap: 0.92, fs: 14 });
s.addNotes("The problem: given a noisy, blurred voxel volume of one object, estimate the area of its true surface. I set five objectives. First, build shapes whose reference area is accurate to better than one part in ten thousand, so the benchmark itself adds no error. Second, measure how biased the classical estimators are. Third, train a 3D CNN and a PointNet to correct the marching-cubes area, and compare them fairly with two calibrated baselines. Fourth, test on shape families the models never saw. Fifth, say when a learned correction is actually worth it.");

// ---------- 4. theory: voxel faces and binary MC ----------
pres.addSection({ title: "Theory" });
s = content("Theory", "Theory", "Voxel faces overestimate by 50% at every resolution");
const sph = Object.fromEntries(D.sphere.map((r) => [r.N, r]));
rows(s, M, 1.85, 6.2, [
  { h: "Staircase area per unit true area", t: "r(n) = |nx| + |ny| + |nz|, between 1 and √3, depending only on the surface normal." },
  { h: "Averaged over all orientations", t: "E|ni| = 1/2, so the face count is 3/2 of the true area. Refining the grid does not change r." },
  { h: "Binary marching cubes", t: "Vertices sit at edge midpoints, so only a few triangle orientations exist. Local methods are not multigrid convergent (Kenmochi and Klette 2000)." },
], { gap: 1.35 });
stat(s, M, 6.0 - 0.15, 2.9, "+" + sph[128]["face_%"].toFixed(1) + "%", "voxel faces, clean sphere, 128³", C.text2, "stat faces");
stat(s, M + 3.2, 6.0 - 0.15, 2.9, "+" + sph[128]["mc_bin_%"].toFixed(1) + "%", "binary MC, clean sphere, 128³", C.text2, "stat mcbin");
img(s, "face_count_theory.png", 7.25, 1.9, 5.5, "Face-count area per unit true area against the polar angle of the surface normal");
s.addNotes("Why does face counting fail? A flat patch with normal n becomes a staircase whose area is the sum of its projections: absolute nx plus ny plus nz. That is 1 only when the surface lines up with the grid, and up to root 3 on the diagonal. Averaged over random orientations it is exactly 3/2, a 50% overestimate, and this does not depend on voxel size at all. On a clean sphere at 128 cubed I measure plus 50.0%. Binary marching cubes is better but has the same kind of problem: its triangles can only take a few orientations, and it is known not to converge. Its error actually grows slightly with resolution, to 8.8% at 128 cubed.");
// adjust stat y to avoid overlap with rows: rows end ~1.85+3*1.35=5.9

// ---------- 5. theory: curvature ----------
s = content("Theory", "Theory", "Greyscale marching cubes shrinks curved surfaces");
rows(s, M, 1.85, 6.0, [
  { h: "Partial volume helps", t: "Interpolating the greyscale value puts vertices at sub-voxel positions. A flat surface is recovered exactly." },
  { h: "Curvature does not", t: "Blurring spreads a curved object outwards, so the half-intensity surface sits slightly inside the true one. The area comes out a little too small." },
  { h: "Prediction for a sphere", t: "Error ≈ -2σ²/R², with R the radius and σ² the total blur (0.44 voxel², or 1.44 after extra smoothing)." },
], { gap: 1.35 });
const trow = (cells, hdr) => cells.map((c) => ({ text: String(c), options: { bold: !!hdr, color: hdr ? "FFFFFF" : "1B2333",
                                                  fill: { color: hdr ? HEX.dk2 : "FFFFFF" }, align: "center", fontSize: 14 } }));
const tbl = [trow(["Grid", "R (vox)", "MC predicted", "MC measured", "Smoothed pred.", "Smoothed meas."], true)];
D.sphere.forEach((r) => tbl.push(trow([`${r.N}³`, r.R_vox.toFixed(1), r["mc_grey_pred_%"].toFixed(2) + "%", r["mc_grey_%"].toFixed(2) + "%",
                                       r["mc_smooth_pred_%"].toFixed(2) + "%", r["mc_smooth_%"].toFixed(2) + "%"])));
s.addTable(tbl, { x: 6.85, y: 2.0, w: 5.9, colW: [0.8, 0.85, 1.05, 1.05, 1.1, 1.05], border: { type: "solid", pt: 0.5, color: "C9CED8" },
                  rowH: 0.45, fontFace: "Calibri", objectName: "curvature table" });
s.addText("Clean spheres, object filling 90% of the grid. The simple curvature formula predicts the measured error closely wherever curvature dominates.",
          { x: 6.85, y: 4.45, w: 5.9, h: 0.8, fontSize: 12, italic: true, color: "52514E", margin: 0, isTextBox: true, objectName: "table caption" });
s.addNotes("Greyscale marching cubes uses the partial-volume information, so a flat surface comes out exactly right. But a curved surface shrinks. Blurring spreads the object's intensity outwards, so the half-intensity surface ends up slightly inside the real one, and more so where the surface bends sharply. For a sphere this gives an error of about minus 2 sigma squared over R squared, where sigma squared is the total blur. The table compares this prediction with what I measured. At 32 cubed: predicted minus 0.56%, measured minus 0.62%. After smoothing: predicted minus 1.82%, measured minus 2.00%. So the negative bias of marching cubes on small objects is not random: it is curvature.");

// ---------- 6. theory: noise and ground truth ----------
s = content("Theory", "Theory", "Noise adds area; exact references make it measurable");
rows(s, M, 1.85, 6.0, [
  { h: "Spurious specks", t: "Each voxel crosses the iso-level by chance with p = Q(0.5/σn). At σn = 0.15, p = 4.3×10⁻⁴: about 112 specks on a 64³ grid." },
  { h: "Rough surfaces", t: "Noise perturbs the surface; a perturbed surface always has more area. This bias is positive." },
  { h: "Two biases, opposite signs", t: "Curvature pulls the area down, noise pushes it up. One fixed correction factor cannot fix both." },
], { gap: 1.35 });
card(s, 6.85, 1.85, 5.9, 4.5, "ground truth card");
s.addText([{ text: "Exact reference area", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } },
           { text: "Exact formulas: sphere, cylinder, torus, ellipsoid.", options: { bullet: true, breakLine: true } },
           { text: "Blobs and superellipsoids: very fine triangle mesh at two levels of detail, combined to cancel the main error. Error below 0.01%.", options: { bullet: true, breakLine: true } },
           { text: "Not marching cubes at 512³ as the spec suggested: binary MC never converges, so it would bake its own bias into the ground truth.", options: { bullet: true, breakLine: true } },
           { text: "14 of 14 pipeline checks pass (e.g. the mesh method and the exact ellipsoid formula agree to 2.5×10⁻⁸).", options: { bullet: true } }],
          { x: 7.15, y: 2.05, w: 5.3, h: 4.1, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true,
            paraSpaceAfter: 6, objectName: "ground truth text" });
s.addNotes("Noise does two things. Background voxels randomly cross the threshold; at noise 0.15 the chance is 4.3 times ten to the minus 4 per voxel, which predicts about 112 tiny specks on a 64 cubed grid, and I measure Euler numbers of about 110. Noise also roughens the surface, and a rough surface always has more area. So curvature pushes the area down and noise pushes it up. To measure all this, the reference area must be exact. I use exact formulas where they exist, and for the blobs a very fine triangle mesh at two levels of detail, combined so that the main error cancels. I deliberately did not use marching cubes at 512 cubed as the reference, because binary marching cubes never converges, so it would put its own bias into the ground truth.");

// ---------- 7. previous work ----------
pres.addSection({ title: "Previous work" });
s = content("Previous work", "Previous work", "Fixed estimators, or learning without ground truth");
const cols = [
  { h: "Local estimators", t: ["Optimal weights for 2×2×2 configurations: Mullikin and Verbeek 1993; Windreich et al. 2003; Lindblad 2005.",
                               "Triangulated surfaces beat voxel surfaces: Nyström et al. 2002."] },
  { h: "Convergence theory", t: ["Binary marching cubes is not multigrid convergent, even for spheres: Kenmochi and Klette 2000.",
                                 "No local digital surface-area estimator converges on convex polytopes: Svane 2013."] },
  { h: "Learning on 3D data", t: ["3D CNNs on voxels (VoxNet, 2015) and PointNet on point sets (2017).",
                                  "DeePore (2020): CNN predicts porous-media properties from micro-CT, without exact ground truth."] },
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
card(s, M, 5.7, W - 2 * M, 1.05, "gap card", "FDF3DC");
s.addText([{ text: "Gap: ", options: { bold: true, color: C.text2 } },
           { text: "no study we found measures how much a learned model reduces surface-area bias against an exact reference, separately by resolution, noise and unseen shape family, and against a calibrated feature-based baseline." }],
          { x: M + 0.3, y: 5.78, w: W - 2 * M - 0.6, h: 0.9, fontSize: 15, color: C.text1, valign: "middle", margin: 0,
            isTextBox: true, objectName: "gap text" });
s.addNotes("Previous work falls into three groups. Digital geometry derived fixed weights for local voxel configurations, for example Lindblad 2005, and showed that triangulated surfaces beat voxel surfaces. Convergence theory showed that binary marching cubes does not converge to the true area even for a sphere, and that no local estimator converges on convex polytopes. And on the learning side, 3D CNNs and PointNet are standard, and DeePore used a CNN to predict porous-media properties from micro-CT, but without an exact ground truth for surface area. The gap I address: measuring, against an exact reference, how much a learned model actually helps, split by resolution, noise and unseen shapes, and compared with a strong simple baseline.");

// ---------- 8. method: data ----------
pres.addSection({ title: "Method" });
s = content("Method", "Method: data", "2,400 synthetic shapes rendered as 14,400 CT-like volumes");
rows(s, M, 1.85, 6.6, [
  { h: "Seven shape families", t: "Train: sphere, ellipsoid, cylinder, torus, smooth random blob (2,000 shapes). Held out: superellipsoid, rough blob (200 each)." },
  { h: "CT forward model", t: "Partial-volume occupancy (5×5×5 supersampling), Gaussian PSF σ = 0.6 voxel, Gaussian noise σn = 0, 0.05, 0.15." },
  { h: "Two resolutions", t: "Every shape at 32³ and 64³ with the same rotation and size: diameters 9.8 to 60 voxels." },
  { h: "Leak-free split", t: "70/15/15 by shape, so all six variants of a shape stay in one split. Fixed seed 51." },
], { gap: 1.2, fs: 14 });
img(s, "example_slices.png", 7.6, 1.85, 5.15, "Central slices of one blob at 32 and 64 cubed and three noise levels");
s.addNotes("The dataset: seven shape families. Five are used for training, about two thousand shapes. Two are held out completely: superellipsoids, which have sharp edges and tips, and rough blobs with fine surface texture. Each shape is voxelized with a simple CT model: partial-volume averaging by supersampling, a Gaussian blur for the point-spread function, and Gaussian noise at three levels. Every shape is rendered at 32 and 64 cubed, so 6 volumes per shape and 14,400 in total. The split is done by shape, never by volume, so the test set has no leakage.");

// ---------- 9. method: models ----------
s = content("Method", "Method: estimators and models", "Learning a correction on top of marching cubes");
const flow = [["CT volume", "N³ greyscale"], ["Greyscale MC", "area A_MC"], ["Model", "predicts y"], ["Corrected area", "Â = A_MC · e^ŷ"]];
flow.forEach((f, i) => {
  const x = M + i * 3.1;
  card(s, x, 1.85, 2.6, 1.1, `flow ${i + 1}`, i === 2 ? "FDF3DC" : HEX.lt2);
  s.addText([{ text: f[0], options: { bold: true, color: C.text2, breakLine: true } }, { text: f[1], options: { color: C.text1 } }],
            { x: x + 0.15, y: 1.92, w: 2.3, h: 0.96, fontSize: 15, align: "center", valign: "middle", margin: 0, isTextBox: true,
              objectName: `flow text ${i + 1}` });
  if (i < 3) s.addShape(pres.shapes.RIGHT_TRIANGLE, { x: x + 2.72, y: 2.25, w: 0.28, h: 0.3, rotate: 90,
                                                     fill: { color: "9AA5B8" }, line: { color: "9AA5B8" }, objectName: `flow arrow ${i + 1}` });
});
s.addText("Target y = log(A_true / A_MC): dimensionless, rotation invariant, so the 48 cube symmetries are exact augmentations.",
          { x: M, y: 3.05, w: W - 2 * M, h: 0.4, fontSize: 14, italic: true, color: "52514E", margin: 0, isTextBox: true, objectName: "target note" });
const models = [
  { h: "3D CNN", t: "5 conv blocks (16 to 256 channels), global pooling. 3.59M parameters.", c: MC.cnn },
  { h: "PointNet", t: "2,048 points from the MC mesh, shared MLP to 1024, max-pool. 0.74M parameters.", c: MC.pointnet },
  { h: "Random Forest", t: "10 simple measurements: MC area ratios, Euler number, noise estimate, shape.", c: MC.rf },
];
models.forEach((m, i) => {
  const x = M + i * 4.13;
  card(s, x, 3.65, 3.88, 2.2, `model card ${i + 1}`);
  s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 3.85, w: 0.32, h: 0.32, fill: { color: m.c }, line: { color: m.c }, objectName: `model dot ${i + 1}` });
  s.addText(m.h, { x: x + 0.7, y: 3.82, w: 3.0, h: 0.4, fontSize: 16, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `model head ${i + 1}` });
  s.addText(m.t, { x: x + 0.25, y: 4.35, w: 3.4, h: 1.4, fontSize: 13, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: `model text ${i + 1}` });
});
s.addText("Baselines: voxel faces, MC on binary / greyscale / smoothed volumes, and MC + a lookup factor per resolution and noise level. Training: Huber loss, AdamW, one-cycle LR, up to 40 epochs, early stopping, Apple-silicon GPU (MPS).",
          { x: M, y: 6.05, w: W - 2 * M, h: 0.7, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, objectName: "baseline note" });
s.addNotes("All learned models predict the same thing: the log of the ratio between the true area and the greyscale marching-cubes area. The final estimate is the marching-cubes area times e to the prediction, so the model is a bias-correction layer on top of the classical method. Because this target does not change under rotation, I can augment exactly with the 48 symmetries of a cube. I compare three learned models: a 3D CNN on the volume, a PointNet on 2,048 points from the mesh, and a Random Forest on ten simple measurements of the volume. The baselines include a simple lookup table with one factor per resolution and noise level.");

// ---------- 10. results: main ----------
pres.addSection({ title: "Results" });
const red32 = Math.floor(D.test.mc_grey["32"] / D.test.cnn["32"]), red64 = Math.floor(D.test.mc_grey["64"] / D.test.cnn["64"]);
s = content("Results", "Results", `Learned corrections cut the error ${red32} to ${red64} times`);
const order = ["mc_grey", "mc_smooth", "lookup", "pointnet", "rf", "cnn"];
s.addChart(pres.charts.BAR, [
  { name: "32³", labels: order.map((m) => NAME[m]), values: order.map((m) => D.test[m]["32"]) },
  { name: "64³", labels: order.map((m) => NAME[m]), values: order.map((m) => D.test[m]["64"]) },
], { x: M, y: 1.75, w: 7.6, h: 4.9, barDir: "bar", barGrouping: "clustered", chartColors: ["9AA5B8", HEX.dk2],
     showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", showLegend: true, legendPos: "b",
     showTitle: true, title: "Test error (MAPE, %), shapes never seen in training", valGridLine: { color: "E4E7EC", size: 0.5 },
     catGridLine: { style: "none" }, valAxisMaxVal: 16, catAxisOrientation: "maxMin", barGapWidthPct: 60, ...chartText,
     objectName: "main results chart" });
stat(s, 8.55, 1.85, 4.2, pct(D.test.cnn["32"]), `3D CNN at 32³ (MC greyscale: ${pct(D.test.mc_grey["32"])})`, MC.cnn, "stat cnn");
stat(s, 8.55, 3.5, 4.2, pct(D.test.rf["64"]), `Random Forest at 64³ (MC greyscale: ${pct(D.test.mc_grey["64"])})`, "B07700", "stat rf");
s.addText([{ text: "Not shown: voxel faces (" + D.test.face["32"].toFixed(0) + "-" + D.test.face["64"].toFixed(0) + "%) and binary MC (" +
                   D.test.mc_bin["32"].toFixed(0) + "-" + D.test.mc_bin["64"].toFixed(0) + "%).", options: { breakLine: true } },
           { text: `PointNet is weaker (${pct(D.test.pointnet["32"])}, ${pct(D.test.pointnet["64"])}): the noisy mesh loses information the volume still has.` }],
          { x: 8.55, y: 5.15, w: 4.2, h: 1.6, fontSize: 13, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6,
            isTextBox: true, objectName: "results notes" });
s.addNotes(`Here are the main results on test shapes the models never saw. Greyscale marching cubes is off by ${pct(D.test.mc_grey["32"])} at 32 cubed and ${pct(D.test.mc_grey["64"])} at 64 cubed. Voxel faces and binary marching cubes are far worse and are left off the chart. A simple lookup table already halves the error. The 3D CNN brings it down to ${pct(D.test.cnn["32"])} and ${pct(D.test.cnn["64"])}. An honest finding: the Random Forest on ten hand-crafted features is just as good at 32 cubed and even better at 64 cubed, ${pct(D.test.rf["64"])}. PointNet is clearly weaker, because sampling points from an already noisy mesh throws information away.`);

// ---------- 11. results: resolution ----------
s = content("Results", "Results: resolution sweep", "Binary MC never converges; greyscale MC does, until noise");
img(s, "bias_vs_resolution.png", (W - 10.6) / 2, 1.65, 10.6, "Mean signed error against grid size for binary, smoothed and greyscale marching cubes at three noise levels");
s.addText([{ text: "Binary MC: bias grows with resolution, to +34% at 128³ under noise (more voxels, more specks).", options: { bullet: true, breakLine: true } },
           { text: "Greyscale MC: -23% at 16³ (curvature), near zero at 128³ when clean, +9.7% at 128³ with σn = 0.15.", options: { bullet: true, breakLine: true } },
           { text: "Learned models sit near zero at both training sizes and all noise levels.", options: { bullet: true } }],
          { x: M, y: 6.15, w: W - 2 * M, h: 0.9, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: "sweep notes" });
s.addNotes("This sweep renders the same 48 shapes at grid sizes from 16 to 128. Binary marching cubes, in green, gets worse as resolution increases, up to plus 34% at 128 cubed with noise, exactly because more voxels means more noise specks. Greyscale marching cubes, in blue, is strongly negative at low resolution, which is the curvature effect, converges on clean data, but is pushed up to plus 10% by noise. Smoothing, in pink, kills the noise but makes small objects much too small. The learned models, the squares, sit close to zero everywhere they were trained.");

// ---------- 12. results: noise and size ----------
s = content("Results", "Results: noise and object size", "Robust to noise; the CNN is best on small objects");
const nl = ["σ = 0", "σ = 0.05", "σ = 0.15"], nm = ["mc_grey", "pointnet", "rf", "cnn"];
[32, 64].forEach((N, i) => {
  s.addChart(pres.charts.LINE, nm.map((m) => ({ name: NAME[m], labels: nl, values: [0.0, 0.05, 0.15].map((v) => D.noise[m][`${N}_${v === 0 ? "0.0" : v}`]) })),
    { x: M + i * 4.2, y: 1.75, w: 4.0, h: 4.0, chartColors: nm.map((m) => MC[m]), lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      showTitle: true, title: `${N}³ grid, test MAPE (%)`, showLegend: i === 1, legendPos: "b", valGridLine: { color: "E4E7EC", size: 0.5 },
      catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 10, ...chartText, objectName: `noise chart ${N}` });
});
if (true) {
  const smallBest = "cnn";
  stat(s, 9.2, 1.85, 3.55, pct(D.small[smallBest]), `${NAME[smallBest]} on objects under 16 voxels (MC greyscale ${pct(D.small.mc_grey)}, RF ${pct(D.small.rf)})`, MC[smallBest], "stat small");
  stat(s, 9.2, 3.55, 3.55, pct(D.noise.cnn["64_0.15"]), `3D CNN at 64³, σn = 0.15 (MC greyscale ${pct(D.noise.mc_grey["64_0.15"])})`, MC.cnn, "stat noise");
}
s.addText("Learned models barely notice noise. Object size matters more than grid size; smoothing is harmful for small objects.",
          { x: M, y: 6.0, w: 8.2, h: 0.7, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: "noise note" });
s.addNotes(`Splitting by noise: greyscale marching cubes jumps to ${pct(D.noise.mc_grey["64_0.15"])} at high noise on the 64 cubed grid, while the CNN only goes from ${pct(D.noise.cnn["64_0.0"])} to ${pct(D.noise.cnn["64_0.15"])}. So the learned models are almost insensitive to noise. Object size matters even more than grid size. For objects smaller than 16 voxels across, marching cubes is off by ${pct(D.small.mc_grey)}, and this is where the CNN is clearly ahead: ${pct(D.small.cnn)}, against ${pct(D.small.rf)} for the Random Forest.`);

// ---------- 13. results: generalisation and failure ----------
s = content("Results", "Results: unseen shapes", "Transfers to unseen shapes; fails below two voxels");
img(s, "heldout_families.png", M, 1.75, 7.2, "MAPE on the test split and the two held-out families at 32 and 64 cubed");
rows(s, 8.2, 1.85, 4.55, [
  { h: "Superellipsoids (unseen)", t: `CNN ${pct(D.super.cnn["32"])} / ${pct(D.super.cnn["64"])}, RF ${pct(D.super.rf["32"])} / ${pct(D.super.rf["64"])}. MC greyscale ${pct(D.super.mc_grey["32"])} / ${pct(D.super.mc_grey["64"])}.` },
  { h: "Rough blobs (unseen)", t: `Hardest: CNN ${pct(D.rough.cnn["32"])} at 32³, biased low. Sub-voxel roughness cannot be recovered.` },
  { h: "Failure case", t: "A torus with a 1.77-voxel tube: MC recovers 42 of 142 voxel². Every model is off by 47 to 66%." },
], { gap: 1.55, fs: 14, labels: ["✓", "!", "✕"], fill: HEX.dk2 });
s.addNotes(`On the two families the models never saw, the error stays far below marching cubes: about ${pct(D.super.cnn["32"])} for the CNN on superellipsoids at 32 cubed, against ${pct(D.super.mc_grey["32"])}. Rough blobs are the hardest case. The models under-predict there, which makes sense: the training shapes are smooth, and detail smaller than a voxel simply is not in the data. And there is one clear failure: a torus whose tube is only 1.77 voxels thick. After blurring, most of the tube never reaches the iso-level, so marching cubes finds only 42 of 142 units of area, and no multiplicative correction can fix a surface that is missing from the mesh.`);

// ---------- 14. conclusion ----------
pres.addSection({ title: "Conclusion" });
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("CONCLUSION", { placeholder: "kicker" });
s.addText("A learned correction removes most of the bias, and theory explains what remains", { placeholder: "title" });
const concl = [
  `Classical estimators are biased in predictable ways: +50% (faces), non-convergent (binary MC), curvature and noise (greyscale MC).`,
  `The best corrections reach ${pct(D.test.cnn["32"])} (32³) and ${pct(D.test.rf["64"])} (64³) on unseen test shapes, from ${pct(D.test.mc_grey["32"])} and ${pct(D.test.mc_grey["64"])}.`,
  `A Random Forest on ten simple measurements matches the 3D CNN; the CNN is worth it mainly for small objects.`,
  `Limits: structures thinner than about two voxels and sub-voxel roughness.`,
];
s.addText(concl.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < concl.length - 1 } })),
          { x: M, y: 3.55, w: 12.1, h: 3.2, fontSize: 18, color: "E6ECF7", valign: "top", margin: 0, paraSpaceAfter: 10,
            isTextBox: true, objectName: "conclusion list" });
s.addNotes("To conclude. The classical estimators are biased in ways we can predict from theory. A learned correction on top of marching cubes removes most of that bias, down to well under 1% on unseen test shapes, and it is robust to noise. A Random Forest on ten simple measurements does just as well as the 3D CNN overall, which is an important honest result. The CNN is the better choice for small objects and coarse grids. The remaining failures, structures thinner than about two voxels and sub-voxel roughness, are exactly where the information is no longer in the image.");

// ---------- 15. limitations ----------
s = content("Conclusion", "Limitations", "What this study does not show yet");
const lims = [
  ["Simplified physics", "White Gaussian noise in the image domain; no projection-domain Poisson noise, beam hardening or rings."],
  ["One object per volume", "Porous media and lattices have many touching structures; untested."],
  ["Fixed acquisition", "Two grid sizes, three noise levels, fixed blur. Outside this range is untested."],
  ["Sub-voxel detail", "Corrections on rough surfaces are a learned prior, not a measurement."],
  ["Limited shapes", "Star-shaped or genus one; no thin sheets or many-handled shapes."],
  ["No tuning, one seed", "64³ CNNs trained only 15 epochs on a laptop; no real CT scan yet."],
];
lims.forEach((l, i) => {
  const x = M + (i % 3) * 4.1, y = 1.85 + Math.floor(i / 3) * 2.45;
  card(s, x, y, 3.85, 2.2, `limitation ${i + 1}`);
  s.addText([{ text: l[0], options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
             { text: l[1], options: { fontSize: 14, color: C.text1 } }],
            { x: x + 0.3, y: y + 0.2, w: 3.25, h: 1.85, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true,
              objectName: `limitation text ${i + 1}` });
});
s.addNotes("The limitations. The CT physics is simplified: noise is added in the image, not in the projections, and there are no beam-hardening or ring artefacts. Every volume has one isolated object, while real porous media have many touching structures. The models only know two grid sizes and three noise levels. On rough surfaces the correction is a learned prior, not a measurement. The shapes are fairly simple. And there was no hyperparameter tuning, the 64 cubed CNNs got only 15 epochs on a laptop, and there is no test on a real CT scan yet. These are the natural next steps.");

// ---------- 16. thanks ----------
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
s.addText("THANK YOU", { placeholder: "kicker" });
s.addText("Questions?", { placeholder: "title" });
s.addText([{ text: "Code, trained weights, data and the preprint article are included in the submission.", options: { breakLine: true } },
           { text: "Manas Jain  |  24123021  |  Department of Physics, IIT Roorkee" }], { placeholder: "body", w: 10 });
s.addNotes("Thank you. The code, trained weights, the dataset and the full article are all part of the submission.");

(async () => {
  const out = path.join("slides", "Project51_ManasJain.pptx");
  await pres.writeFile({ fileName: out });
  if (applyTheme) await applyTheme(out, THEME);
  console.log("wrote", out);
})();
