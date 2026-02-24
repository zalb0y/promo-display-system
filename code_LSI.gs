// ============================================================
// CODE.GS — LSI Version
// Bound script di Google Slides LSI
// Deploy as Web App: Execute as ME, Access: Anyone
// ============================================================

const STORE_TYPE = 'LSI';
const MAX_PRODUCTS_PER_SLIDE = 9;

const REGION_ORDER   = { 'Region 1': 1, 'Region 2': 2, 'Region 3': 3 };
const DIVISION_ORDER = { 'Dry Food': 1, 'Meal Solution': 2, 'Fresh Food': 3, 'H&B HOME CARE': 4, 'Non Food': 5 };

// Division card colors (RGB 0-1)
const DIVISION_CARD_COLORS = {
  'Dry Food':       { red: 0.98, green: 0.78, blue: 0.35 },
  'Fresh Food':     { red: 0.55, green: 0.78, blue: 0.55 },
  'Meal Solution':  { red: 0.72, green: 0.25, blue: 0.30 },
  'H&B HOME CARE':  { red: 0.68, green: 0.80, blue: 0.92 },
  'Non Food':       { red: 0.45, green: 0.55, blue: 0.65 }
};

// ==========================
// WEB APP
// ==========================
function doGet() {
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.storeType = STORE_TYPE;
  return tmpl.evaluate()
    .setTitle('Lotte Mart Promo — ' + STORE_TYPE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================
// SUBMIT PRODUCT
// ==========================
function submitProductWithImage(formData) {
  try {
    let imageUrl = '';

    // Upload image if provided
    if (formData.imageBase64) {
      const res = uploadImage(
        formData.imageBase64,
        'promo_' + formData.storeName.replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now() + '.png'
      );
      if (res.success) imageUrl = res.url;
      else return { success: false, message: 'Gagal upload gambar: ' + res.message };
    }

    // Open THIS presentation (bound script)
    const presentation = SlidesApp.getActivePresentation();

    const region    = formData.region;
    const storeName = formData.storeName;
    const division  = formData.division;
    const category  = formData.category;
    const prodNm    = formData.prodNm;
    const stk       = formData.stk;
    const productKey = storeName + '|' + division + '|' + prodNm;

    // Find or create slide & slot
    const result = findOrCreateSlide(presentation, region, storeName, division, productKey, category, prodNm);

    // Place product
    placeProduct(result.slide, result.slotIndex, category, prodNm, stk, imageUrl, division);

    return {
      success: true,
      message: '"' + prodNm + '" berhasil ditambahkan di slot ' + (result.slotIndex + 1) + '!'
    };
  } catch (err) {
    Logger.log('Error: ' + err);
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

// ==========================
// FIND OR CREATE SLIDE
// ==========================
function findOrCreateSlide(presentation, region, storeName, division, productKey, category, prodNm) {
  const slides = presentation.getSlides();

  // --- Pass 1: cari product yang sudah ada (overwrite/timpa) ---
  for (let i = 1; i < slides.length; i++) {
    const meta = parseMeta(slides[i]);
    if (!meta || meta.type === 'region_separator') continue;
    if (meta.storeName === storeName && meta.division === division && meta.products) {
      for (let p = 0; p < meta.products.length; p++) {
        if (meta.products[p] && meta.products[p].key === productKey) {
          return { slide: slides[i], slotIndex: p };
        }
      }
    }
  }

  // --- Pass 2: cari slide dengan slot kosong ---
  const matching = [];
  for (let i = 1; i < slides.length; i++) {
    const meta = parseMeta(slides[i]);
    if (!meta || meta.type === 'region_separator') continue;
    if (meta.storeName === storeName && meta.division === division) {
      matching.push({ slide: slides[i], meta: meta, index: i });
    }
  }

  if (matching.length > 0) {
    const last = matching[matching.length - 1];
    const products = last.meta.products || newProductArray();
    for (let s = 0; s < MAX_PRODUCTS_PER_SLIDE; s++) {
      if (!products[s]) {
        products[s] = { key: productKey, prodNm: prodNm, category: category };
        last.meta.products = products;
        setSlideNotes(last.slide, JSON.stringify(last.meta));
        return { slide: last.slide, slotIndex: s };
      }
    }

    // Semua slot penuh → buat slide lanjutan
    const newSlide = createProductSlide(presentation, region, storeName, division, last.index + 1);
    const prods = newProductArray();
    prods[0] = { key: productKey, prodNm: prodNm, category: category };
    saveMeta(newSlide, region, storeName, division, prods);
    return { slide: newSlide, slotIndex: 0 };
  }

  // --- Pass 3: slide baru — insert di posisi yang benar ---
  maybeInsertRegionSeparator(presentation, region);
  const insertIdx = findInsertPosition(presentation, region, storeName, division);
  const newSlide = createProductSlide(presentation, region, storeName, division, insertIdx);
  const prods = newProductArray();
  prods[0] = { key: productKey, prodNm: prodNm, category: category };
  saveMeta(newSlide, region, storeName, division, prods);
  return { slide: newSlide, slotIndex: 0 };
}

// ==========================
// POSITION — urut region → store → division
// ==========================
function findInsertPosition(presentation, region, storeName, division) {
  const slides = presentation.getSlides();
  const rOrd = REGION_ORDER[region] || 99;
  const dOrd = DIVISION_ORDER[division] || 99;

  for (let i = 1; i < slides.length; i++) {
    const m = parseMeta(slides[i]);
    if (!m) continue;
    const mR = REGION_ORDER[m.region] || 99;

    if (m.type === 'region_separator') {
      if (mR > rOrd) return i;
      continue;
    }

    const mD = DIVISION_ORDER[m.division] || 99;
    if (mR > rOrd) return i;
    if (mR === rOrd) {
      if (m.storeName > storeName) return i;
      if (m.storeName === storeName && mD > dOrd) return i;
    }
  }
  return slides.length;
}

// ==========================
// REGION SEPARATOR
// ==========================
function maybeInsertRegionSeparator(presentation, region) {
  const slides = presentation.getSlides();

  // Cek apakah separator sudah ada
  for (let i = 1; i < slides.length; i++) {
    const m = parseMeta(slides[i]);
    if (m && m.type === 'region_separator' && m.region === region) return;
  }

  // Cek apakah sudah ada konten lain
  let hasContent = false;
  for (let i = 1; i < slides.length; i++) {
    const m = parseMeta(slides[i]);
    if (m && m.type !== 'region_separator') { hasContent = true; break; }
  }

  if (hasContent || region !== 'Region 1') {
    const idx = findSeparatorPos(presentation, region);
    const slide = insertBlankSlide(presentation, idx);
    slide.getBackground().setSolidFill('#CC0000');

    const tb1 = slide.insertTextBox(region.toUpperCase(), 100, 150, 520, 100);
    tb1.getText().getTextStyle().setFontSize(48).setBold(true).setForegroundColor('#FFFFFF').setFontFamily('Arial');

    const tb2 = slide.insertTextBox('LOTTE MART — Area Display Promo', 100, 260, 520, 50);
    tb2.getText().getTextStyle().setFontSize(18).setForegroundColor('#FFFFFF').setFontFamily('Arial');

    setSlideNotes(slide, JSON.stringify({ type: 'region_separator', region: region }));
  }
}

function findSeparatorPos(presentation, region) {
  const slides = presentation.getSlides();
  const rOrd = REGION_ORDER[region] || 99;
  for (let i = 1; i < slides.length; i++) {
    const m = parseMeta(slides[i]);
    if (!m) continue;
    if ((REGION_ORDER[m.region] || 99) >= rOrd) return i;
  }
  return slides.length;
}

// ==========================
// CREATE PRODUCT SLIDE
// ==========================
function createProductSlide(presentation, region, storeName, division, insertIndex) {
  const slide = insertBlankSlide(presentation, insertIndex);
  slide.getBackground().setSolidFill('#FFFFFF');

  const cc = DIVISION_CARD_COLORS[division] || { red: 0.5, green: 0.5, blue: 0.5 };

  // --- HEADER ---
  // Logo
  const logo = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 20, 10, 50, 50);
  logo.getFill().setSolidFill('#CC0000');
  logo.getText().setText('L');
  logo.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor('#FFFFFF');
  logo.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  addText(slide, STORE_TYPE, 75, 8, 150, 18, 10, '#CC0000', false);
  addText(slide, storeName, 75, 24, 280, 22, 13, '#333333', true);
  addText(slide, region, 75, 44, 150, 18, 10, '#CC0000', false);

  const divBox = addText(slide, division, 280, 18, 200, 30, 18, '#333333', true);
  divBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const lbl = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 590, 15, 120, 35);
  lbl.getFill().setSolidFill('#F5F5F0');
  lbl.getBorder().setWeight(0.5).getLineFill().setSolidFill('#CCCCCC');
  lbl.getText().setText('Area Display Promo');
  lbl.getText().getTextStyle().setFontSize(9).setBold(true).setForegroundColor('#333333').setFontFamily('Arial');
  lbl.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // --- 3×3 PRODUCT CARDS ---
  const sx = 18, sy = 72, cw = 228, ch = 120, gx = 12, gy = 10;
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      const idx = row * 3 + col;
      const x = sx + col * (cw + gx);
      const y = sy + row * (ch + gy);

      const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, cw, ch);
      card.getFill().setSolidFill(r255(cc.red), r255(cc.green), r255(cc.blue));
      card.getBorder().setWeight(0).setTransparent();
      card.setDescription('CARD_BG_' + idx);

      const catTb = slide.insertTextBox('', x + 5, y + 5, 100, 14);
      catTb.getText().getTextStyle().setFontSize(8).setBold(true).setForegroundColor('#333333').setFontFamily('Arial');
      catTb.setDescription('CATEGORY_' + idx);

      addText(slide, 'Prod_nm', x + 5, y + 22, 50, 12, 7, '#666666', false);

      const pBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 5, y + 34, 80, 18);
      pBox.getFill().setSolidFill('#FFFFFF');
      pBox.getBorder().setWeight(0.5).getLineFill().setSolidFill('#DDDDDD');
      pBox.getText().getTextStyle().setFontSize(7).setForegroundColor('#333333').setFontFamily('Arial');
      pBox.setDescription('PROD_NM_' + idx);

      addText(slide, 'Stk', x + 5, y + 56, 30, 12, 7, '#666666', false);

      const sBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 5, y + 68, 60, 18);
      sBox.getFill().setSolidFill('#FFFFFF');
      sBox.getBorder().setWeight(0.5).getLineFill().setSolidFill('#DDDDDD');
      sBox.getText().getTextStyle().setFontSize(7).setForegroundColor('#333333').setFontFamily('Arial');
      sBox.setDescription('STK_' + idx);

      const img = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 95, y + 10, 125, 100);
      img.getFill().setSolidFill('#FFFFFF');
      img.getBorder().setWeight(0.5).setDashStyle(SlidesApp.DashStyle.DASH).getLineFill().setSolidFill('#AAAAAA');
      img.setDescription('IMG_PLACEHOLDER_' + idx);
    }
  }

  return slide;
}

// ==========================
// PLACE PRODUCT
// ==========================
function placeProduct(slide, slotIndex, category, prodNm, stk, imageUrl, division) {
  const els = slide.getPageElements();
  for (let i = 0; i < els.length; i++) {
    const d = els[i].getDescription();
    if (d === 'CATEGORY_' + slotIndex)  els[i].asShape().getText().setText(category);
    if (d === 'PROD_NM_' + slotIndex)   els[i].asShape().getText().setText(prodNm);
    if (d === 'STK_' + slotIndex)        els[i].asShape().getText().setText(String(stk));

    if ((d === 'IMG_PLACEHOLDER_' + slotIndex || d === 'IMG_' + slotIndex) && imageUrl) {
      const l = els[i].getLeft(), t = els[i].getTop(), w = els[i].getWidth(), h = els[i].getHeight();
      els[i].remove();
      try {
        slide.insertImage(imageUrl, l, t, w, h).setDescription('IMG_' + slotIndex);
      } catch (e) { Logger.log('Img err: ' + e); }
    }
  }

  // Update metadata
  try {
    const meta = JSON.parse(getSlideNotes(slide));
    if (!meta.products) meta.products = newProductArray();
    meta.products[slotIndex] = { key: meta.storeName + '|' + meta.division + '|' + prodNm, prodNm: prodNm, category: category };
    setSlideNotes(slide, JSON.stringify(meta));
  } catch (e) {}
}

// ==========================
// IMAGE → DRIVE
// ==========================
function uploadImage(base64Data, fileName) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', fileName);
    const folders = DriveApp.getFoldersByName('LotteMart_PromoImages');
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('LotteMart_PromoImages');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://drive.google.com/uc?export=view&id=' + file.getId() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================
// HELPERS
// ==========================
function parseMeta(slide) {
  const n = getSlideNotes(slide);
  if (!n) return null;
  try { return JSON.parse(n); } catch (e) { return null; }
}

function saveMeta(slide, region, storeName, division, products) {
  setSlideNotes(slide, JSON.stringify({
    storeType: STORE_TYPE, region: region,
    storeName: storeName, division: division,
    products: products
  }));
}

function newProductArray() { return new Array(MAX_PRODUCTS_PER_SLIDE).fill(null); }

function getSlideNotes(slide) {
  try {
    const s = slide.getNotesPage().getPlaceholder(SlidesApp.PlaceholderType.BODY);
    return s ? s.getText().asString().trim() : '';
  } catch (e) { return ''; }
}

function setSlideNotes(slide, text) {
  try {
    slide.getNotesPage().getPlaceholder(SlidesApp.PlaceholderType.BODY).getText().setText(text);
  } catch (e) {}
}

function insertBlankSlide(presentation, index) {
  const slides = presentation.getSlides();
  return index >= slides.length
    ? presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK)
    : presentation.insertSlide(index, SlidesApp.PredefinedLayout.BLANK);
}

function addText(slide, text, x, y, w, h, size, color, bold) {
  const tb = slide.insertTextBox(text, x, y, w, h);
  const ts = tb.getText().getTextStyle();
  ts.setFontSize(size).setForegroundColor(color).setFontFamily('Arial');
  if (bold) ts.setBold(true);
  return tb;
}

function r255(v) { return Math.round(v * 255); }
