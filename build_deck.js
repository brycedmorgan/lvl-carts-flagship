const pptxgen = require("pptxgenjs");
const path = require("path");

const ASSETS = path.join(__dirname, "assets");

// Brand palette — Desert Highline (Stitch design system)
const C = {
  void:     "0B0B0B",  // primary background
  void2:    "131411",  // secondary background
  void3:    "1C1C19",  // card background
  offWhite: "F4F1EC",
  gold:     "D4AF37",  // primary accent
  goldSoft: "E9C349",
  tan:      "B89B6E",
  muted:    "8E9192",
  outline:  "444748"
};

const F = { head: "Georgia", body: "Calibri" };

let pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";  // 13.3 × 7.5
pres.author = "LVL Carts Pitch";
pres.title  = "LVL Carts — Digital Flagship Proposal";

const W = 13.333, H = 7.5;

/* ---------- helpers ---------- */
function bgVoid(s) { s.background = { color: C.void }; }
function goldHairline(s, x, y, w) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.015, fill: { color: C.gold }, line: { type: "none" } });
}
function tanRule(s, x, y, w) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.01, fill: { color: C.tan, transparency: 60 }, line: { type: "none" } });
}
function pageNum(s, n, total) {
  s.addText(`${String(n).padStart(2,"0")} / ${String(total).padStart(2,"0")}`,
    { x: W-1.6, y: H-0.5, w: 1.2, h: 0.3, fontFace: F.body, fontSize: 9,
      color: C.tan, align: "right", charSpacing: 3 });
}
function brandMark(s, x=0.5, y=0.5) {
  s.addText("LVL", { x, y, w: 0.7, h: 0.4, fontFace: F.head, fontSize: 22, color: C.gold, charSpacing: 2, bold: true, margin: 0 });
  s.addText("CARTS", { x: x+0.78, y: y+0.05, w: 1.4, h: 0.3, fontFace: F.body, fontSize: 12, color: C.offWhite, charSpacing: 6, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x, y: y+0.42, w: 1.85, h: 0.012, fill: { color: C.gold }, line: { type: "none" } });
}
function slidePreamble(s, kicker) {
  brandMark(s);
  s.addText(kicker, { x: W-3.5, y: 0.55, w: 3, h: 0.3, fontFace: F.body, fontSize: 9, color: C.tan, charSpacing: 4, align: "right", margin: 0 });
  goldHairline(s, 0.5, 1.05, W-1);
}

const TOTAL = 12;
let n = 2;  // cover is 01, content slides start at 02

/* ============ SLIDE 1: COVER ============ */
{
  let s = pres.addSlide(); bgVoid(s);

  // diagonal hairline accent
  goldHairline(s, 0.5, 1.6, 4.5);

  s.addText("LVL", { x: 0.5, y: 0.5, w: 1.0, h: 0.5, fontFace: F.head, fontSize: 28, color: C.gold, bold: true, charSpacing: 3, margin: 0 });
  s.addText("CARTS", { x: 1.55, y: 0.6, w: 2.5, h: 0.4, fontFace: F.body, fontSize: 14, color: C.offWhite, charSpacing: 8, margin: 0 });

  s.addText("DIGITAL FLAGSHIP", { x: 0.5, y: 1.85, w: 6, h: 0.35, fontFace: F.body, fontSize: 11, color: C.tan, charSpacing: 6, margin: 0 });

  s.addText("A new website for\nSt. George's premier\ngolf cart destination.", {
    x: 0.5, y: 2.35, w: 11, h: 3.0, fontFace: F.head, fontSize: 60, color: C.offWhite, italic: false, charSpacing: -1
  });

  goldHairline(s, 0.5, 6.0, 4.5);
  s.addText("Prepared for the LVL Carts founding family · St. George, Utah", {
    x: 0.5, y: 6.15, w: 9, h: 0.3, fontFace: F.body, fontSize: 11, color: C.muted
  });
  s.addText("MAY 2026", { x: W-2.0, y: 6.15, w: 1.5, h: 0.3, fontFace: F.body, fontSize: 11, color: C.gold, charSpacing: 4, align: "right" });
}

/* ============ SLIDE 2: THE OPPORTUNITY ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "01 · THE OPPORTUNITY");

  s.addText("Your inventory is luxury.\nYour website isn't (yet).", {
    x: 0.5, y: 1.45, w: 12, h: 1.6, fontFace: F.head, fontSize: 40, color: C.offWhite, charSpacing: -1
  });

  s.addText("LVL Carts sells $10K – $13K vehicles with curated finishes, custom builds, and white-glove delivery — but the current site reads like a generic dealer page running on a stock template. The shopping experience and the showroom experience don't match.", {
    x: 0.5, y: 3.15, w: 8, h: 2.0, fontFace: F.body, fontSize: 16, color: C.offWhite, paraSpaceAfter: 8
  });

  // Right-side stat block
  s.addShape(pres.shapes.RECTANGLE, { x: 9.5, y: 1.45, w: 0.04, h: 4.5, fill: { color: C.gold }, line: { type: "none" } });
  s.addText("WHY IT MATTERS", { x: 9.7, y: 1.45, w: 3.2, h: 0.35, fontFace: F.body, fontSize: 10, color: C.tan, charSpacing: 4 });
  s.addText("Conversion rate on luxury e-com sites is 3–5× generic templates. Buyers researching $13K carts don't comparison-shop on price — they comparison-shop on trust signals: photography, story, polish.", {
    x: 9.7, y: 1.85, w: 3.4, h: 4.0, fontFace: F.body, fontSize: 13, color: C.offWhite, paraSpaceAfter: 6
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 3: WHAT WE'RE PROPOSING ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "02 · THE PROPOSAL");

  s.addText("Four moves. One brand.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 40, color: C.offWhite, charSpacing: -1
  });
  s.addText("A complete digital flagship — designed for the way you actually sell.", {
    x: 0.5, y: 2.4, w: 12, h: 0.5, fontFace: F.body, fontSize: 16, color: C.tan
  });

  const cards = [
    { kicker: "01", title: "Redesign", body: "Premium black + gold aesthetic that mirrors your showroom. Built on the Stitch \"Desert Highline\" design system." },
    { kicker: "02", title: "Rides Rental Sync", body: "Keep your existing inventory backend. Pipe live cart data into the new site via Rides Rental's platform integrations." },
    { kicker: "03", title: "AI Concierge", body: "An always-on chatbot that answers inventory, financing, delivery, and scheduling questions — and books test drives." },
    { kicker: "04", title: "SEO & Local",  body: "Schema markup, local-business signals, fast page speed, and content built for \"St. George golf carts\" searches." }
  ];

  const startY = 3.2;
  cards.forEach((c, i) => {
    const x = 0.5 + i * 3.13;
    s.addShape(pres.shapes.RECTANGLE, { x, y: startY, w: 0.04, h: 3.4, fill: { color: C.gold }, line: { type: "none" } });
    s.addText(c.kicker, { x: x+0.2, y: startY, w: 1, h: 0.35, fontFace: F.body, fontSize: 10, color: C.tan, charSpacing: 4 });
    s.addText(c.title,  { x: x+0.2, y: startY+0.4, w: 2.85, h: 0.55, fontFace: F.head, fontSize: 22, color: C.offWhite, bold: false });
    s.addText(c.body,   { x: x+0.2, y: startY+1.05, w: 2.85, h: 2.3, fontFace: F.body, fontSize: 12, color: C.offWhite, paraSpaceAfter: 4 });
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 4: BEFORE → AFTER ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "03 · BEFORE → AFTER");

  s.addText("From template to flagship.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 40, color: C.offWhite, charSpacing: -1
  });

  // Two columns
  s.addText("CURRENT", { x: 0.5, y: 2.5, w: 5.5, h: 0.35, fontFace: F.body, fontSize: 11, color: C.muted, charSpacing: 4 });
  s.addText("PROPOSED", { x: 7.0, y: 2.5, w: 5.8, h: 0.35, fontFace: F.body, fontSize: 11, color: C.gold,  charSpacing: 4 });
  tanRule(s, 0.5, 2.9, 5.5);
  goldHairline(s, 7.0, 2.9, 5.8);

  const before = [
    "Generic Rides Rental template",
    "Lorem-ipsum testimonials still live",
    "\"Venom EV\" still in meta tags (rebranded to Royal EV)",
    "Stock orange/black palette — same as 1,000 other dealers",
    "No financing flow on-site",
    "No chatbot — every question = a phone call",
    "Thin SEO. No schema. No local pages."
  ];
  const after = [
    "Desert Highline aesthetic — black, gold, tan",
    "Real testimonials from Southern Utah customers",
    "Brand-consistent across every touchpoint",
    "Editorial photography, big serif headlines, sharp UI",
    "Inline financing form — get your out-the-door price",
    "AI assistant: inventory, financing, delivery, scheduling",
    "AutoDealer schema + LocalBusiness markup + city pages"
  ];

  before.forEach((t,i)=>{
    s.addText("·", { x: 0.5, y: 3.1 + i*0.42, w: 0.2, h: 0.35, fontFace: F.body, fontSize: 16, color: C.muted });
    s.addText(t,   { x: 0.7, y: 3.1 + i*0.42, w: 5.4, h: 0.4,  fontFace: F.body, fontSize: 13, color: C.offWhite });
  });
  after.forEach((t,i)=>{
    s.addText("→", { x: 7.0, y: 3.1 + i*0.42, w: 0.3, h: 0.35, fontFace: F.body, fontSize: 13, color: C.gold });
    s.addText(t,   { x: 7.35, y: 3.1 + i*0.42, w: 5.5, h: 0.4, fontFace: F.body, fontSize: 13, color: C.offWhite });
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 5: THE NEW SITE (DESIGN PREVIEW) ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "04 · DESIGN PREVIEW");

  s.addText("Your real inventory. New frame.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 36, color: C.offWhite, charSpacing: -1
  });
  s.addText("Below: actual product photos from your current site, dropped into the new design language. Week 3 includes a full retouching pass — clean backgrounds, color consistency, editorial cropping.", {
    x: 0.5, y: 2.4, w: 11, h: 0.95, fontFace: F.body, fontSize: 13, color: C.tan
  });

  // 3 inventory thumbnails as faux-screens
  try {
    s.addImage({ path: path.join(ASSETS, "cart-pearl-white.jpg"),  x: 0.5, y: 3.4, w: 4.0, h: 3.0, sizing: { type: "cover", w: 4.0, h: 3.0 } });
    s.addImage({ path: path.join(ASSETS, "cart-tiffany-blue.jpg"), x: 4.65, y: 3.4, w: 4.0, h: 3.0, sizing: { type: "cover", w: 4.0, h: 3.0 } });
    s.addImage({ path: path.join(ASSETS, "cart-royal-white.jpg"),  x: 8.8, y: 3.4, w: 4.0, h: 3.0, sizing: { type: "cover", w: 4.0, h: 3.0 } });
  } catch(e) { /* ignore */ }

  // captions
  ["ATLAS · 4-PASS · LIFTED  $12,990","ATLAS · TIFFANY BLUE  $12,990","ROYAL EV · CROWN 4  $9,990"]
    .forEach((c,i)=>{
      s.addText(c, { x: 0.5 + i*4.15, y: 6.5, w: 4.0, h: 0.3, fontFace: F.body, fontSize: 9, color: C.gold, charSpacing: 4 });
    });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 6: RIDES RENTAL INTEGRATION ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "05 · RIDES RENTAL INTEGRATION");

  s.addText("Keep your backend. Replace the front-end.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 36, color: C.offWhite, charSpacing: -1
  });
  s.addText("You already use Rides Rental Software for inventory, leads, and billing. We don't replace that. We replace the customer-facing site — and pipe the data through.", {
    x: 0.5, y: 2.45, w: 8.5, h: 0.9, fontFace: F.body, fontSize: 14, color: C.tan
  });

  // Diagram-ish row
  const rowY = 3.9;
  function box(x, label, sub, accent) {
    s.addShape(pres.shapes.RECTANGLE, { x, y: rowY, w: 3.0, h: 1.8, fill: { color: C.void2 }, line: { color: accent || C.outline, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: rowY, w: 0.05, h: 1.8, fill: { color: accent || C.gold }, line: { type: "none" } });
    s.addText(label, { x: x+0.2, y: rowY+0.2, w: 2.7, h: 0.4, fontFace: F.head, fontSize: 18, color: C.offWhite });
    s.addText(sub,   { x: x+0.2, y: rowY+0.7, w: 2.7, h: 1.0, fontFace: F.body, fontSize: 11, color: C.tan, paraSpaceAfter: 3 });
  }
  function arrow(x) {
    s.addText("→", { x, y: rowY + 0.6, w: 0.5, h: 0.6, fontFace: F.body, fontSize: 28, color: C.gold, align: "center" });
  }
  box(0.5,  "Rides Rental",   "Inventory · Leads · Customer DB · Payments · Reports", C.gold);
  arrow(3.6);
  box(4.15, "Sync Layer",     "Webhooks + nightly sync. Inventory + leads flow both ways.", C.tan);
  arrow(7.25);
  box(7.8,  "New LVL Site",   "Stitch design. Live inventory. Chatbot. Forms.", C.gold);
  // standalone services
  s.addText("INTEGRATIONS RIDES RENTAL ALREADY SUPPORTS", { x: 0.5, y: 5.95, w: 8, h: 0.3, fontFace: F.body, fontSize: 10, color: C.muted, charSpacing: 4 });
  s.addText("Stripe · Authorize.net · PayPal · Lightspeed · Dealertrack DMS · Google Analytics · Zapier · YouTube",
    { x: 0.5, y: 6.25, w: 12, h: 0.5, fontFace: F.body, fontSize: 12, color: C.offWhite });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 7: AI CHATBOT ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "06 · THE AI CONCIERGE");

  s.addText("An assistant that knows your inventory cold.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 34, color: C.offWhite, charSpacing: -1
  });
  s.addText("Available 24/7. Trained on your products, pricing, financing, delivery, and hours.", {
    x: 0.5, y: 2.4, w: 8.5, h: 0.5, fontFace: F.body, fontSize: 14, color: C.tan
  });

  // 3 capability cards
  const caps = [
    { t: "Inventory", b: "Lists what's on the floor, by brand, by passenger count, by color. Pulls live from Rides Rental." },
    { t: "Education", b: "Atlas vs Royal EV, lithium vs lead-acid, range, charging, LSV street-legal package, warranty." },
    { t: "Logistics",  b: "Delivery zones, financing pre-qual, scheduling test drives, hours, address — and books appointments." }
  ];
  caps.forEach((c,i)=>{
    const x = 0.5 + i*4.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.3, w: 4.0, h: 2.5, fill: { color: C.void2 }, line: { color: C.outline, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 3.3, w: 0.05, h: 2.5, fill: { color: C.gold }, line: { type: "none" } });
    s.addText(c.t, { x: x+0.25, y: 3.45, w: 3.7, h: 0.5, fontFace: F.head, fontSize: 22, color: C.offWhite });
    s.addText(c.b, { x: x+0.25, y: 4.0,  w: 3.6, h: 1.7, fontFace: F.body, fontSize: 12, color: C.offWhite });
  });

  // Stat row
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 6.05, w: 12.3, h: 0.7, fill: { color: C.void2 }, line: { type: "none" } });
  s.addText("EXPECTED IMPACT", { x: 0.7, y: 6.15, w: 2.5, h: 0.5, fontFace: F.body, fontSize: 10, color: C.tan, charSpacing: 4 });
  s.addText("+30% lead capture · –40% phone-tag · 100% off-hours coverage",
    { x: 4, y: 6.15, w: 8.7, h: 0.5, fontFace: F.body, fontSize: 13, color: C.gold, align: "right" });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 8: SEO PLAN ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "07 · SEO & LOCAL");

  s.addText("Own \"St. George golf carts.\"", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 38, color: C.offWhite, charSpacing: -1
  });
  s.addText("Most local dealers don't bother with structured SEO. We'll fix that in week one.",
    { x: 0.5, y: 2.45, w: 11, h: 0.5, fontFace: F.body, fontSize: 14, color: C.tan });

  const items = [
    { t: "Schema.org markup",   b: "AutoDealer + LocalBusiness JSON-LD on every page so Google reads inventory, hours, address." },
    { t: "Meta + OG tags",      b: "Per-page titles, meta descriptions, Open Graph, Twitter cards. Already sloppy on the live site." },
    { t: "Local landing pages", b: "/golf-carts-st-george, /golf-carts-washington-ut, /golf-carts-hurricane — each ranking for its own city." },
    { t: "Page speed",          b: "Single-file build, lazy-loaded images, Tailwind CDN. Lighthouse 95+ on every page." },
    { t: "Content engine",      b: "Quarterly long-form posts (\"Atlas vs Club Car\", \"Best golf carts for desert terrain\") to capture top-of-funnel." },
    { t: "Google Business sync", b: "Hours, photos, posts, reviews — kept in lockstep with the site so Google trusts you." }
  ];

  items.forEach((it,i)=>{
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 6.3;
    const y = 3.3 + row * 1.3;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.04, h: 1.1, fill: { color: C.gold }, line: { type: "none" } });
    s.addText(it.t, { x: x+0.2, y: y, w: 5.9, h: 0.4, fontFace: F.head, fontSize: 18, color: C.offWhite });
    s.addText(it.b, { x: x+0.2, y: y+0.45, w: 5.9, h: 0.7, fontFace: F.body, fontSize: 12, color: C.tan });
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 9: TIMELINE ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "08 · TIMELINE");

  s.addText("Live in 4 weeks.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 40, color: C.offWhite, charSpacing: -1
  });

  const phases = [
    { w: "WEEK 1", t: "Discovery + Brand Lock", b: "Confirm voice. Photo audit. Real-customer outreach for testimonials. Final brand colors." },
    { w: "WEEK 2", t: "Build Phase",            b: "Site build complete. Rides Rental sync wired. Chatbot trained on knowledge base." },
    { w: "WEEK 3", t: "Content + SEO",          b: "Local landing pages. Schema. Meta. Photo retouching. GBP optimization." },
    { w: "WEEK 4", t: "Launch + Handoff",       b: "QA across devices. DNS cutover. Analytics + Search Console set up. Owner training." }
  ];

  phases.forEach((p,i)=>{
    const x = 0.5 + i*3.13;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.9, w: 3.0, h: 0.05, fill: { color: C.gold }, line: { type: "none" } });
    s.addText(p.w, { x, y: 3.1, w: 3.0, h: 0.35, fontFace: F.body, fontSize: 11, color: C.gold, charSpacing: 4 });
    s.addText(p.t, { x, y: 3.5, w: 3.0, h: 0.5, fontFace: F.head, fontSize: 20, color: C.offWhite });
    s.addText(p.b, { x, y: 4.1, w: 3.0, h: 2.5, fontFace: F.body, fontSize: 12, color: C.tan });
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 10: WHAT YOU GET ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "09 · DELIVERABLES");

  s.addText("Everything you need. Nothing you don't.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 36, color: C.offWhite, charSpacing: -1
  });

  const items = [
    "New website — 6+ pages, fully responsive, premium design system",
    "Rides Rental data sync (live inventory + lead capture)",
    "AI chatbot trained on your inventory, financing, delivery, hours",
    "On-page SEO: schema, meta, alt text, local landing pages",
    "Editorial photography retouching pass on existing inventory shots",
    "Google Business Profile optimization + review-flow setup",
    "Analytics + Search Console + heatmap install",
    "30 days post-launch support — bug fixes, content tweaks, training",
    "Source files (HTML, CSS, JS, image library) — you own everything"
  ];

  items.forEach((it,i)=>{
    const y = 2.7 + i*0.45;
    s.addText("✓", { x: 0.5, y: y, w: 0.4, h: 0.35, fontFace: F.body, fontSize: 16, color: C.gold, bold: true });
    s.addText(it,  { x: 0.95, y: y, w: 11.8, h: 0.4, fontFace: F.body, fontSize: 14, color: C.offWhite });
  });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 11: THE DEAL ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "10 · THE DEAL");

  s.addText("Our ask is unusual.", {
    x: 0.5, y: 1.45, w: 12, h: 0.9, fontFace: F.head, fontSize: 40, color: C.offWhite, charSpacing: -1
  });
  s.addText("We believe in this build. We want skin in the game — and we'd rather drive your product than send an invoice.",
    { x: 0.5, y: 2.45, w: 11, h: 0.7, fontFace: F.body, fontSize: 16, color: C.tan });

  // Big offer card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 12.3, h: 3.4, fill: { color: C.void2 }, line: { color: C.gold, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 0.06, h: 3.4, fill: { color: C.gold }, line: { type: "none" } });

  s.addText("THE OFFER", { x: 0.85, y: 3.5, w: 4, h: 0.4, fontFace: F.body, fontSize: 11, color: C.tan, charSpacing: 4 });
  s.addText("Full digital flagship build, traded for one Royal EV Crown 4 ($9,990 retail).", {
    x: 0.85, y: 3.95, w: 11.5, h: 1.3, fontFace: F.head, fontSize: 28, color: C.offWhite, charSpacing: -1
  });

  s.addText("WHAT WE PROVIDE", { x: 0.85, y: 5.4, w: 5, h: 0.3, fontFace: F.body, fontSize: 10, color: C.gold, charSpacing: 4 });
  s.addText("Site · Chatbot · SEO · Rides Rental sync · 30 days support",
    { x: 0.85, y: 5.7, w: 6, h: 0.5, fontFace: F.body, fontSize: 13, color: C.offWhite });

  s.addText("WHAT YOU PROVIDE", { x: 7.0, y: 5.4, w: 5, h: 0.3, fontFace: F.body, fontSize: 10, color: C.gold, charSpacing: 4 });
  s.addText("One Royal EV Crown 4 + a testimonial if you love the work",
    { x: 7.0, y: 5.7, w: 6, h: 0.5, fontFace: F.body, fontSize: 13, color: C.offWhite });

  pageNum(s, n++, TOTAL);
}

/* ============ SLIDE 12: NEXT STEPS ============ */
{
  let s = pres.addSlide(); bgVoid(s);
  slidePreamble(s, "11 · NEXT STEPS");
  pageNum(s, 12, TOTAL);

  s.addText("Let's build it.", {
    x: 0.5, y: 2.0, w: 12, h: 1.4, fontFace: F.head, fontSize: 80, color: C.offWhite, charSpacing: -2
  });
  goldHairline(s, 0.5, 3.7, 4.5);
  s.addText("If the offer makes sense, we can kick off discovery this week and have a live demo URL in your inbox in 7 days.",
    { x: 0.5, y: 3.95, w: 11, h: 1.0, fontFace: F.body, fontSize: 18, color: C.tan });

  // Footer info
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.7, w: 12.3, h: 0.012, fill: { color: C.tan, transparency: 60 }, line: { type: "none" } });
  s.addText("CONTACT", { x: 0.5, y: 5.9, w: 4, h: 0.3, fontFace: F.body, fontSize: 10, color: C.gold, charSpacing: 4 });
  s.addText("Bryce Morgan · brycedmorgan@gmail.com", { x: 0.5, y: 6.2, w: 6, h: 0.4, fontFace: F.body, fontSize: 14, color: C.offWhite });
  s.addText("LVL CARTS", { x: W-3.5, y: 5.9, w: 3, h: 0.3, fontFace: F.body, fontSize: 10, color: C.gold, charSpacing: 4, align: "right" });
  s.addText("(435) 414-1180  ·  info@lvlcarts.com  ·  677 N 3050 E, St. George UT", { x: W-7, y: 6.2, w: 6.5, h: 0.4, fontFace: F.body, fontSize: 12, color: C.tan, align: "right" });
}

pres.writeFile({ fileName: path.join(__dirname, "LVL_Carts_Digital_Flagship_Pitch.pptx") })
  .then(f => console.log("Saved:", f));
