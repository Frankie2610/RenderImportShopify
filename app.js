const ONE_XLSX_PATH = "./ONE.xlsx";

/* ===== LOAD TXT DESCRIPTIONS ===== */
let DESCRIPTION_POOL = [];

function hashString(str = "") {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = ((hash << 5) - hash) + str.charCodeAt(i);
        hash |= 0;
    }
    return Math.abs(hash);
}

function pickDescriptionBySKU(sku = "") {
    if (!DESCRIPTION_POOL.length) return "";
    const idx = hashString(sku) % DESCRIPTION_POOL.length;
    return DESCRIPTION_POOL[idx];
}

function stripDiacritics(s = "") {
    // dùng NFD + loại combining marks
    return s && s.normalize ? s.normalize("NFD").replace(/[\u0300-\u036f]/g, "") : String(s || "");
}

function normalizeKey(key = "") {
    if (key === null || typeof key === "undefined") return "";
    return String(key)
        .replace(/\u00A0/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function normalizeVendor(vendor = "") {
    if (vendor === null || typeof vendor === "undefined") return "";
    return String(vendor)
        .normalize("NFKD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\u00A0/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function normalizeGenderKey(gender = "") {
    const g = normalizeVietnamese(gender);

    if (g.includes("nam") || g === "male") return "male";
    if (g.includes("nữ") || g.includes("nu") || g === "female") return "female";
    return "unisex"; // fallback an toàn
}

const descriptionPools = {
    watches: watchDescriptions,
    watch: watchDescriptions,
    dongho: watchDescriptions,

    sunglasses: sunglassesDescriptions,
    kính: sunglassesDescriptions,
    kinhram: sunglassesDescriptions
};

function pickDescription({
    productType,
    sku,
    gender,
    titleRaw
}) {
    const pool = descriptionPools[(productType || "").toLowerCase()];
    if (!pool) return "";

    const g = normalizeGenderKey(gender);
    const list = pool[g] || pool.unisex;
    if (!list || !list.length) return "";

    const index = hashString(sku || titleRaw) % list.length;

    return list[index].replace("{{NAME}}", `<b>${titleRaw}</b>`);
}

// extractValue giữ nguyên logic (nhưng an toàn hơn với kiểu non-string)
function extractValue(text = "", label) {
    if (!text || typeof text !== "string") return "";

    const normalize = str => (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    const normLabel = normalize(label);
    const lines = text.split(/\r?\n/);

    for (let line of lines) {
        const parts = line.split(/[:：]/);
        if (parts.length >= 2) {
            const key = normalize(parts[0]);
            if (key.includes(normLabel)) {
                return parts.slice(1).join(":").trim();
            }
        }
    }
    return "";
}

// Chuẩn hóa tiêu đề thành in hoa (bản gốc của bạn)
function normalizeKeys(row) {
    const normalized = {};
    for (const key in row) {
        const cleanKey = normalizeKey(key);
        const cleanValue = typeof row[key] === "string" ? row[key].trim() : row[key];
        normalized[cleanKey] = cleanValue;
    }
    return normalized;
}

// Hàm capitalizeWords (giữ lại)
function capitalizeWords(str = "") {
    return String(str || "")
        .toLowerCase()
        .split(/\s+/)
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
}

// normalizeVietnamese (giữ lại)
function normalizeVietnamese(str = "") {
    return String(str || "")
        .normalize("NFC")
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();
}

// extractStrapMaterial (giữ logic của bạn, nhưng trả kết quả dạng "Dây ..." như bạn mong)
function extractStrapMaterial(input = "") {
    if (!input) return "";
    let s = normalizeVietnamese(input);

    s = s
        .replace(/\bdây\b/g, "")
        .replace(/\bband\b/g, "")
        .replace(/\bmàu\b/g, "")
        .trim();

    const materialMap = [
        { keys: ["nhựa", "resin"], value: "nhựa" },
        { keys: ["cao su", "rubber"], value: "cao su" },
        { keys: ["da bò", "genuine leather"], value: "da bò" },
        { keys: ["da", "leather"], value: "da" },
        { keys: ["kim loại", "inox", "thép", "metal", "steel"], value: "kim loại" },
        { keys: ["lưới", "mesh"], value: "lưới" },
        { keys: ["vải", "dệt", "canvas"], value: "vải" },
        { keys: ["nylon"], value: "nylon" },
        { keys: ["silicone"], value: "silicone" },
        { keys: ["gốm", "ceramic"], value: "gốm" },
        { keys: ["mạ vàng", "gold plated", "mạ", "mạ gold"], value: "kim loại" },
    ];

    for (const item of materialMap) {
        for (const key of item.keys) {
            if (s.includes(key)) {
                const result = item.value.charAt(0) + item.value.slice(1);
                return "Dây " + result;
            }
        }
    }
    return "";
}

// extractStrapColor (giữ logic)
function extractStrapColor(input = "") {
    if (!input) return "";

    let s = normalizeVietnamese(input);
    s = s
        .replace(/\bdây\b/g, "")
        .replace(/\bband\b/g, "")
        .replace(/\bstrap\b/g, "")
        .replace(/\bcó\b/g, "")
        .replace(/\bcó\b/g, "")
        .trim();

    const colorMap = [
        { keys: ["đen", "black"], value: "Đen" },
        { keys: ["trắng", "white"], value: "Trắng" },
        { keys: ["nâu", "brown"], value: "Nâu" },
        { keys: ["đỏ", "red"], value: "Đỏ" },
        { keys: ["bạc", "silver"], value: "Bạc" },
        { keys: ["kem", "beige"], value: "Kem" },
        { keys: ["xám", "ghi", "grey"], value: "Xám" },
        { keys: ["xanh dương", "blue"], value: "Xanh Dương" },
        { keys: ["xanh navy", "navy"], value: "Xanh Navy" },
        { keys: ["xanh lá", "green"], value: "Xanh" },
        { keys: ["xanh"], value: "Xanh" },
        { keys: ["vàng gold", "gold"], value: "Vàng Gold" },
        { keys: ["vàng hồng", "rose gold"], value: "Vàng Hồng" },
        { keys: ["vàng"], value: "Vàng" },
        { keys: ["hồng", "pink"], value: "Hồng" },
    ];

    for (const item of colorMap) {
        for (const key of item.keys) {
            if (s.includes(key)) return item.value;
        }
    }
    return "";
}

// extractCaseColor (giữ logic)
function extractCaseColor(color = "") {
    if (!color) return "";
    let s = normalizeVietnamese(color);
    s = s.replace(/\bcó\b/g, "").replace(/\bcó\b/g, "").trim();

    const colorMap = [
        { keys: ["đen"], value: "Đen" },
        { keys: ["hồng"], value: "Hồng" },
        { keys: ["trắng"], value: "Trắng" },
        { keys: ["nâu"], value: "Nâu" },
        { keys: ["đỏ"], value: "Đỏ" },
        { keys: ["tím"], value: "Tím" },
        { keys: ["bạc", "silver"], value: "Bạc" },
        { keys: ["kem"], value: "Kem" },
        { keys: ["ivory"], value: "Ivory" },
        { keys: ["champagne"], value: "Champagne" },
        { keys: ["xám", "ghi"], value: "Xám" },
        { keys: ["xanh dương", "blue"], value: "Xanh Dương" },
        { keys: ["xanh navy", "navy"], value: "Xanh Navy" },
        { keys: ["xanh lá", "xanh rêu"], value: "Xanh lá" },
        { keys: ["xanh"], value: "Xanh" },
        { keys: ["vàng gold", "gold"], value: "Vàng Gold" },
        { keys: ["vàng hồng", "rose gold"], value: "Vàng Hồng" },
        { keys: ["vàng"], value: "Vàng" },
        { keys: ["hoạ tiết", "pattern"], value: "Hoạ Tiết" },
        { keys: ["hai màu", "two tone", "2 tone"], value: "Hai Màu" },
        { keys: ["điện tử"], value: "Điện tử" },
    ];

    for (const item of colorMap) {
        for (const key of item.keys) {
            if (s.includes(key)) {
                return item.value;
            }
        }
    }
    return "";
}

// BRAND_COUNTRY mapping (chuyển hóa keys như bạn làm)
const RAW_BRAND_COUNTRY = {
    "Alexander McQueen": "Anh",
    "Dolce & Gabbana": "Ý",
    "Balenciaga": "Tây Ban Nha",
    "Chopard": "Thụy Sĩ",
    "Montblanc": "Thụy Sĩ",
    "Miu Miu": "Ý",
    "Ferragamo": "Thụy Sĩ",
    "Ted Baker": "Anh Quốc",
    "Philipp Plein": "Đức",
    "Guess": "Hoa Kỳ",
    "Adidas": "Đức",
    "Furla": "Ý",
    "Locman": "Ý",
    "Missoni": "Ý",
    "Versace": "Ý",
    "Vivienne Westwood": "Anh",
    "By Far": "Bungari",
    "Burberry": "Anh Quốc",
    "Fendi": "Ý",
    "Jimmy Choo": "Anh Quốc",
    "Roberto Cavalli": "Ý",
    "Givenchy": "Pháp",
    "Stella McCartney": "Anh Quốc",
    "Versus By Versace": "Ý"
};
const BRAND_COUNTRY = Object.fromEntries(
    Object.entries(RAW_BRAND_COUNTRY).map(([k, v]) => [
        stripDiacritics(k).replace(/\u00A0/g, "").replace(/\s+/g, " ").trim().toUpperCase(),
        v
    ])
);

// PRODUCT_CATEGORY_MAP + SEO_TITLE_MAP (giữ lại)
const PRODUCT_CATEGORY_MAP = {
    Watches: "Apparel & Accessories > Jewelry > Watches",
    Sunglasses: "Apparel & Accessories > Clothing Accessories > Sunglasses",
    Handbags: "Apparel & Accessories > Handbags, Wallets & Cases > Handbags",
    Earrings: "Apparel & Accessories > Jewelry > Earrings",
    Necklaces: "Apparel & Accessories > Jewelry > Necklaces",
    Bracelets: "Apparel & Accessories > Jewelry > Bracelets",
    Rings: "Apparel & Accessories > Jewelry > Rings",
    Wallets: "Apparel & Accessories > Handbags, Wallets & Cases > Wallets & Money Clips > Wallets",
    HairPin: "Health & Beauty > Personal Care > Hair Care > Hair Styling Tool Accessories > Hair Curler Clips & Pins",
    WatchBand: "Apparel & Accessories > Jewelry > Watch Bands"
};

const SEO_TITLE_MAP = {
    Watches: (gender, vendor, title) => `Đồng Hồ ${gender} ${vendor} ${title}`,
    Sunglasses: (gender, vendor, title) => `Gọng Kính${gender} ${vendor} ${title}`,
    Handbags: (gender, vendor, title) => `Túi Xách ${gender} ${vendor} ${title}`,
    Earrings: (gender, vendor, title) => `Bông Tai ${gender} ${vendor} ${title}`,
    Necklaces: (gender, vendor, title) => `Vòng Cổ ${gender} ${vendor} ${title}`,
    Bracelets: (gender, vendor, title) => `Vòng Tay ${gender} ${vendor} ${title}`,
    Rings: (gender, vendor, title) => `Nhẫn ${gender} ${vendor} ${title}`,
    HairPin: (gender, vendor, title) => `Kẹp Tóc ${gender} ${vendor} ${title}`,
    WatchBand: (gender, vendor, title) => `Dây Đồng Hồ ${gender} ${vendor} ${title}`
};

// generateBodyHTML: gần nguyên gốc, nhưng dùng stripDiacritics/normalizeKey để truy xuất row
function generateBodyHTML(row, gender, vendor, sku, shortDesc, type) {
    const desc = row[normalizeKey("Description")] || "";
    const material = extractValue(desc, "Chất liệu") || extractValue(desc, "Chất liệu vỏ máy") || row[normalizeKey("Chất liệu")];
    const dialColor = extractValue(desc, "Màu sắc") || extractValue(desc, "Màu mặt số") || row[normalizeKey("Màu sắc")] || row[normalizeKey("Color")];
    const color = extractValue(desc, "Màu sắc");
    const watchBezel = extractValue(desc, "Viền đồng hồ");
    const size = extractValue(desc, "Size") || extractValue(desc, "Đường kính") || row[normalizeKey("Size")];
    const glassMaterial = extractValue(desc, "Chất liệu kính") || extractValue(desc, "Chất liệu mặt kính");
    const waterResistant = extractValue(desc, "Chống nước");
    const rawMachine = extractValue(desc, "Máy");
    const machine = ["Quartz", "Automatic"].includes(rawMachine) ? rawMachine : "Quartz";
    const strap = extractValue(desc, "Dây đeo");
    const normalizedVendor = normalizeVendor(vendor || "");
    const country = BRAND_COUNTRY[normalizedVendor] || "chưa xác định";

    return type === "Watches" ? `
<p>${shortDesc || ""}</p>
<p><strong>Thông số sản phẩm</strong></p>
<ul>
<li>Mã SKU: ${sku || ""}</li>
<li>Giới tính: ${gender || ""}</li>
<li>Chất liệu vỏ máy: ${material || ""}</li>
<li>Viền đồng hồ: ${watchBezel || ""}</li>
<li>Đường kính: ${size || ""}</li>
<li>Màu mặt số: ${dialColor || ""}</li>
<li>Chất liệu kính: ${glassMaterial || ""}</li>
<li>Chống nước: ${waterResistant || ""}</li>
<li>Máy: ${machine || ""}</li>
<li>Dây đeo: ${strap || ""}</li>
${(vendor === "FURLA" || vendor === "LOCMAN")
            ? `<li>Xuất xứ thương hiệu: ${country}</li>`
            : (vendor === "FERRAGAMO" || vendor === "VERSACE")
                ? `<li>Xuất xứ thương hiệu: Ý</li>
           <li>Sản xuất tại: Thụy Sĩ</li>`
                : `<li>Xuất xứ thương hiệu: ${country}</li>
           <li>Xuất xứ máy: Máy Nhật</li>`
        }
</ul>`.trim()
        : `
<p>${shortDesc || ""}</p>
<p><strong>Thông số sản phẩm</strong></p>
<ul>
<li>Mã SKU: ${sku || ""}</li>
<li>Giới tính: ${gender || ""}</li>
<li>Chất liệu: ${material || ""}</li>
<li>Màu sắc: ${color || ""}</li>
<li>Size: ${size || ""}</li>
<li>Xuất xứ thương hiệu: ${country}</li>
</ul>`.trim();
}

// --- MAIN: load ONE.xlsx (template) ---
const oneStatusEl = document.getElementById("oneStatus");
let ORIGINAL_ONE_HEADERS = []; // exact header strings from ONE.xlsx
let NORMALIZED_ONE_HEADERS = []; // normalized version of above (using normalizeKey)

async function loadONE() {
    try {
        oneStatusEl.textContent = "đang tải...";
        const res = await fetch(ONE_XLSX_PATH);
        if (!res.ok) throw new Error("Không thể load ONE.xlsx từ " + ONE_XLSX_PATH);
        const ab = await res.arrayBuffer();
        const wbOne = XLSX.read(ab, { type: "array" });
        const firstSheet = wbOne.Sheets[wbOne.SheetNames[0]];
        const headerRow = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })[0] || [];
        ORIGINAL_ONE_HEADERS = headerRow.map(h => String(h || "").trim());
        NORMALIZED_ONE_HEADERS = ORIGINAL_ONE_HEADERS.map(h => normalizeKey(h));
        oneStatusEl.textContent = `loaded (${ORIGINAL_ONE_HEADERS.length} cột)`;
        updateProcessButtonState();
    } catch (err) {
        console.error(err);
        oneStatusEl.textContent = "Lỗi load ONE.xlsx";
    }
}
loadONE();

// UI bindings
const twoFileInput = document.getElementById("twoFile");
const processBtn = document.getElementById("processBtn");
const statusEl = document.getElementById("status");
const downloadLink = document.getElementById("downloadLink");

let twoFile = null;
twoFileInput.addEventListener("change", (e) => {
    twoFile = e.target.files && e.target.files[0];
    updateProcessButtonState();
});

function updateProcessButtonState() {
    processBtn.disabled = !(twoFile && ORIGINAL_ONE_HEADERS.length > 0);
}

// --- Mapping candidates (giữ logic mapping tương tự bản trước, nhưng dùng normalizeKey) ---
const ONE_TO_TWO_CANDIDATES = {
    "TITLE": ["TÊN SẢN PHẨM", "TEN SAN PHAM", "TITLE"],
    "VENDOR": ["THƯƠNG HIỆU", "THUONG HIEU", "VENDOR"],
    "TYPE": ["TYPE", "LOẠI", "LOAI", "CATEGORY"],
    "GENDER": ["GIỚI TÍNH", "GIOI TINH", "GENDER"],
    "VARIANT SKU": ["MÃ SẢN PHẨM", "MA SAN PHAM", "MÃ SKU", "SKU", "MA SKU", "MÃ HÀNG", "MA HANG"],
    "VARIANT INVENTORY QTY": ["SỐ LƯỢNG", "SO LUONG", "QUANTITY"],
    "VARIANT PRICE": ["GIÁ GIẢM", "GIÁ SALE", "GIA GIAM", "PRICE", "PRICE SALE", "GIÁ GIẢM", "GIẢ GIẢM"],
    "VARIANT COMPARE AT PRICE": ["GIÁ BÁN LẺ", "GIA BAN LE", "GIA BAN", "COMPARE AT PRICE", "GIÁ BÁN LẺ"],
    "SEO DESCRIPTION": ["MÔ TẢ", "MO TA", "SHORT DESCRIPTION", "META DESCRIPTION", "MÔ TẢ NGẮN"],
    "BODY (HTML)": ["DESCRIPTION", "Description", "CHI TIET", "THÔNG TIN CHI TIẾT"]
};

// Helper: tìm giá trị đầu tiên không rỗng trong normalized row
function getFromNormalizedRow(nr, candidates = []) {
    for (const c of candidates) {
        const nk = normalizeKey(c);
        if (typeof nr[nk] !== "undefined" && nr[nk] !== null && String(nr[nk]).toString().trim() !== "") {
            return String(nr[nk]).toString().trim();
        }
    }
    return "";
}

// process button click: đọc TWO, map, export
processBtn.addEventListener("click", async () => {
    if (!twoFile) return;
    statusEl.textContent = "Đang đọc TWO.xlsx ...";
    try {
        const ab = await twoFile.arrayBuffer();
        const wbTwo = XLSX.read(ab, { type: "array" });
        const sheetTwo = wbTwo.Sheets[wbTwo.SheetNames[0]];
        const rawDataTwo = XLSX.utils.sheet_to_json(sheetTwo, { defval: "" });
        // normalize each TWO row
        const dataTwo = rawDataTwo.map(r => normalizeKeys(r));

        // Build dataOne (array of objects with normalized keys)
        const dataOne = dataTwo.map((row) => {
            const desc = row[normalizeKey("Description")] || "";
            const vendor = row[normalizeKey("Vendor")] || row[normalizeKey("Thương hiệu")] || row[normalizeKey("THƯƠNG HIỆU")] || "";
            const type = row[normalizeKey("Type")] || row[normalizeKey("TYPE")] || "";
            const sku = (row[normalizeKey("Product code")] || "").toString().replace(/[\r\n\"]+/g, "").trim()
                || row[normalizeKey("Mã hàng")]
                || row[normalizeKey("Mã SKU")]
                || row[normalizeKey("MÃ SẢN PHẨM")]
                || row[normalizeKey("Mã sản phẩm")]
                || "";

            const strapRaw = extractValue(desc, "Dây đeo");
            const strap = extractStrapMaterial(strapRaw);
            const strapColor = extractStrapColor(strapRaw);

            const waterResistant = extractValue(desc, "Chống nước");
            const gender = extractValue(desc, "Giới tính")
                || row[normalizeKey("Giới tính")]
                || row[normalizeKey("Gender")]
                || "";

            const titleRaw = capitalizeWords(
                row[normalizeKey("Tên sản phẩm")] ||
                row[normalizeKey("TEN SAN PHAM")] ||
                row[normalizeKey("TITLE")]
            );

            const title = type === "Watches"
                ? `Đồng Hồ ${gender} ${titleRaw}`
                : titleRaw;
            let shortDesc =
                row[normalizeKey("Mô tả")] ||
                row[normalizeKey("MÔ TẢ")] ||
                "";

            function stripHTML(html = "") {
                return html.replace(/<[^>]+>/g, "");
            }

            if (!shortDesc || !shortDesc.trim()) {
                shortDesc = pickDescription({
                    productType: type,   // Watches / Sunglasses
                    sku,
                    gender,
                    titleRaw
                });

            }
            // SEO DESCRIPTION → loại bỏ <b>
            const seoDescription = stripHTML(shortDesc);

            const salePrice = row[normalizeKey("Giá sale")] || row[normalizeKey("Giá giảm")] || row[normalizeKey("GIÁ GIẢM")] || row[normalizeKey("VARIANT PRICE")] || "";
            const originalPrice = row[normalizeKey("Giá bán lẻ")] || row[normalizeKey("Giá bán")] || row[normalizeKey("GIÁ BÁN LẺ")] || "";
            const qty = row[normalizeKey("Quantity")] || row[normalizeKey("Số lượng")] || row[normalizeKey("SỐ LƯỢNG")] || "";

            const size = extractValue(desc, "Size")
                || extractValue(desc, "Đường kính")
                || extractValue(desc, "Kích thước")
                || row[normalizeKey("Size")] || "";

            const caseColorRaw = extractValue(desc, "Màu mặt số") || row[normalizeKey("PRODUCT.METAFIELDS.CUSTOM.CASECOLOR")] || "";

            const caseColor = extractCaseColor(caseColorRaw);
            const faceShape = extractValue(desc, "Hình dạng mặt số") || row[normalizeKey("PRODUCT.METAFIELDS.CUSTOM.FACESHAPE")] || "";

            // build TAGS (giữ logic bạn có)
            const tags =
                `${gender === "Unisex" ? "Unisex, Nam, Nữ" : gender}, ${vendor}, ${size}, ${type === "Watches" ? "Watch, Watches" : type
                }, ${strap}, ${waterResistant}`;

            // SEO_TITLE like original
            const seoTitle = (SEO_TITLE_MAP[type] || ((gender, vendor, titleRaw) =>
                `${capitalizeWords(vendor)} ${titleRaw}`
            ))(gender, capitalizeWords(vendor), titleRaw);

            const handle = `${vendor}-${sku}`.replace(/\s+/g, "-");

            // prepare object with normalized headers as keys (so mapping later is easy)
            const obj = {};

            // Fill typical normalized keys (use normalizeKey to align with NORMALIZED_ONE_HEADERS)
            obj[normalizeKey("TITLE")] = title || "";
            obj[normalizeKey("VENDOR")] = vendor || "";
            obj[normalizeKey("TYPE")] = type || "";
            obj[normalizeKey("PRODUCT CATEGORY")] = PRODUCT_CATEGORY_MAP[type] || "";
            obj[normalizeKey("GENDER")] = gender || "";
            obj[normalizeKey("TAGS")] = tags || "";
            obj[normalizeKey("VARIANT SKU")] = sku || "";
            obj[normalizeKey("VARIANT INVENTORY QTY")] = qty || "";
            obj[normalizeKey("VARIANT FULFILLMENT SERVICE")] = "manual";
            obj[normalizeKey("STATUS")] = "draft";
            obj[normalizeKey("INVENTORY POLICY")] = "deny";
            obj[normalizeKey("VARIANT INVENTORY TRACKER")] = "shopify";
            obj[normalizeKey("SEO TITLE")] = seoTitle || "";
            obj[normalizeKey("HANDLE")] = handle || "";
            obj[normalizeKey("PUBLISHED")] = "TRUE";
            obj[normalizeKey("VARIANT PRICE")] = salePrice || "";
            obj[normalizeKey("VARIANT COMPARE AT PRICE")] = originalPrice || "";
            obj[normalizeKey("VARIANT REQUIRES SHIPPING")] = "TRUE";
            obj[normalizeKey("VARIANT TAXABLE")] = "TRUE";
            obj[normalizeKey("VARIANT INVENTORY POLICY")] = "DENY";
            obj[normalizeKey("SEO DESCRIPTION")] = seoDescription || "";
            obj[normalizeKey("IMAGE SRC")] = "https://cdn.shopify.com/s/files/1/0862/7906/1824/files/L_M_Logo.jpg?v=1767866894"
            obj[normalizeKey("BODY (HTML)")] = generateBodyHTML(row, gender, vendor, sku, shortDesc, type) || "";

            // metafields
            obj[normalizeKey("MÀU MẶT SỐ (PRODUCT.METAFIELDS.CUSTOM.CASECOLOR)")] = caseColor || "";
            obj[normalizeKey("HÌNH DẠNG MẶT SỐ (PRODUCT.METAFIELDS.CUSTOM.FACESHAPE)")] = faceShape || "Mặt tròn";
            obj[normalizeKey("KÍCH THƯỚC MẶT SỐ (PRODUCT.METAFIELDS.CUSTOM.FACESIZE)")] = size || "";
            obj[normalizeKey("GIỚI TÍNH (PRODUCT.METAFIELDS.CUSTOM.GENDER)")] = gender || "";
            obj[normalizeKey("MÀU DÂY (PRODUCT.METAFIELDS.CUSTOM.M_U_D_Y)")] = strapColor || caseColor || "";
            obj[normalizeKey("CHẤT LIỆU DÂY (PRODUCT.METAFIELDS.CUSTOM.BANDMATERIAL)")] = strap || "";

            // Also keep original normalized row fields so fallback can use them
            Object.keys(row).forEach(k => {
                if (!obj[k]) obj[k] = row[k];
            });

            return obj;
        });

        const exportFormatEl = document.getElementById("exportFormat");
        const exportFormat = exportFormatEl ? exportFormatEl.value : "xlsx";

        /* ===== REORDER DATA theo ONE.xlsx ===== */
        const reordered = dataOne.map(normalizedRow => {
            const newRow = {};
            for (let i = 0; i < ORIGINAL_ONE_HEADERS.length; i++) {
                const origHeader = ORIGINAL_ONE_HEADERS[i];
                const normHeader = NORMALIZED_ONE_HEADERS[i];

                let value = "";
                if (typeof normalizedRow[normHeader] !== "undefined") {
                    value = normalizedRow[normHeader];
                } else if (typeof normalizedRow[origHeader] !== "undefined") {
                    value = normalizedRow[origHeader];
                } else {
                    value = "";
                }
                newRow[origHeader] = value == null ? "" : value;
            }
            return newRow;
        });

        if (!reordered.length) {
            statusEl.textContent = "Không có dữ liệu để xuất file.";
            return;
        }

        /* ===== BUILD WORKSHEET ===== */
        const ws = XLSX.utils.json_to_sheet(reordered, {
            header: ORIGINAL_ONE_HEADERS
        });

        let blob;
        let fileName;

        if (exportFormat === "csv") {
            // ===== EXPORT CSV =====
            const csv = XLSX.utils.sheet_to_csv(ws, { FS: "," });
            blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
            fileName = "FILLED_ONE_FROM_TWO.csv";
        } else {
            // ===== EXPORT XLSX =====
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, ws, "Sheet1");
            const wbout = XLSX.write(newWb, { bookType: "xlsx", type: "array" });
            blob = new Blob([wbout], { type: "application/octet-stream" });
            fileName = "FILLED_ONE_FROM_TWO.xlsx";
        }

        /* ===== DOWNLOAD ===== */
        const url = URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = fileName;
        downloadLink.style.display = "inline-block";
        downloadLink.textContent = `Tải file kết quả (${exportFormat.toUpperCase()})`;

        statusEl.textContent =
            "Hoàn tất — nhấn 'Tải file kết quả' để lưu file.";

    } catch (err) {
        console.error(err);
        statusEl.textContent = "Lỗi khi xử lý file: " + (err && err.message ? err.message : String(err));
    }
});
