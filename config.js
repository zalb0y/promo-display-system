// TODO: Masukkan URL Web App Apps Script setelah di-deploy
const APPS_SCRIPT_URL = "URL_WEB_APP_ANDA_DISINI";

const db = {
    stores: {
        LSI: {
            "Region 1": ["6001 - Pasar Rebo", "6003 - Kelapa Gading", "6007 - Alam Sutera", "6010 - Medan", "6014 - Palembang", "6015 - Pekanbaru", "6021 - Jatake", "6029 - Batam", "6031 - Lampung", "6006 - Ciputat", "6022 - Serang", "6039 - Serpong"],
            "Region 2": ["6004 - Meruya", "6005 - Bandung", "6008 - Cibitung", "6023 - Cikarang", "6024 - Cirebon", "6027 - Tasikmalaya", "6034 - Karawang", "6036 - Padalarang", "6038 - Tegal", "6018 - Bekasi", "6026 - Bogor", "6030 - Pakansari", "6002 - Sidoarjo"],
            "Region 3": ["6009 - Denpasar", "6013 - Makasar", "6017 - Banjarmasin", "6020 - Balikpapan", "6028 - Mastrip", "6032 - Samarinda", "6033 - Manado", "6037 - Mataram", "6011 - Semarang", "6016 - Yogyakarta", "6019 - Solo"]
        },
        LMI: {
            "Region 1": ["04003 - FESTIVAL CITY LINK", "04004 - KELAPA GADING", "04005 - KUNINGAN CITY", "04009 - TAMAN SURYA", "04010 - MEDAN CENTRE POINT", "04021 - GREEN PRAMUKA"],
            "Region 2": ["04001 - GANDARIA", "04006 - BINTARO", "04007 - PANAKUKANG", "04008 - FATMAWATI", "04013 - SOLO BARU", "04020 - PAKUWON MALL SBY"],
            "Region 3": ["04022 - BALI BEER"]
        }
    },
    categories: {
        "Dry Food": ["11 - Biscuit/Snacks", "17 - Bulk Product", "21 - Sauces&Spices", "23 - Drinks", "24 - Milk"],
        "Meal Solution": ["80 - BAKERY", "82 - Delica"],
        "Fresh Food": ["31 - Fish", "32 - Meat", "33 - Fruits", "34 - Vegetables", "35 - Dairy & Frozen"],
        "H&B HOME CARE": ["14 - Home Care", "19 - H&B"],
        "Non Food": ["86 - IT/GADGET", "87 - Small Appliance", "88 - BIG APPLIANCE", "ELC - Electronic", "51 - Kitchen", "57 - Bathroom", "85 - DIY", "13 - Interior & Bedding", "62 - Textile", "71 - Stationary/Toys"]
    }
};

const els = {
    type: document.getElementById('storeType'),
    region: document.getElementById('region'),
    store: document.getElementById('storeName'),
    division: document.getElementById('division'),
    category: document.getElementById('category'),
    form: document.getElementById('promoForm'),
    btn: document.getElementById('submitBtn'),
    loader: document.getElementById('loadingIndicator')
};

function updateStoreOptions() {
    const type = els.type.value;
    const region = els.region.value;
    els.store.innerHTML = '<option value="">-- Pilih Toko --</option>';
    
    if(type) {
        els.region.disabled = false;
        if(region && db.stores[type][region]) {
            els.store.disabled = false;
            db.stores[type][region].forEach(s => els.store.innerHTML += `<option value="${s}">${s}</option>`);
        } else {
            els.store.disabled = true;
        }
    } else {
        els.region.disabled = true;
        els.store.disabled = true;
    }
}

els.type.addEventListener('change', updateStoreOptions);
els.region.addEventListener('change', updateStoreOptions);

els.division.addEventListener('change', function() {
    const div = this.value;
    els.category.innerHTML = '<option value="">-- Pilih Kategori --</option>';
    if(div && db.categories[div]) {
        els.category.disabled = false;
        db.categories[div].forEach(c => els.category.innerHTML += `<option value="${c}">${c}</option>`);
    } else {
        els.category.disabled = true;
    }
});

const fileToBase64 = file => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = error => reject(error);
});

els.form.addEventListener('submit', async (e) => {
    e.preventDefault();
    els.btn.disabled = true;
    els.loader.style.display = 'block';

    try {
        const file = document.getElementById('imageUpload').files[0];
        const base64Img = await fileToBase64(file);

        const payload = {
            storeType: els.type.value,
            region: els.region.value,
            storeName: els.store.value,
            division: els.division.value,
            category: els.category.value,
            prodNm: document.getElementById('prodNm').value,
            stk: document.getElementById('stk').value,
            imgName: file.name,
            imgMime: file.type,
            imgData: base64Img
        };

        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        });
        
        const result = await response.text();
        alert("Sukses: " + result);
        els.form.reset();
        ['region', 'store', 'category'].forEach(id => document.getElementById(id).disabled = true);
        
    } catch (error) {
        alert("Terjadi kesalahan sistem. Cek console log.");
        console.error(error);
    } finally {
        els.btn.disabled = false;
        els.loader.style.display = 'none';
    }
});
