import re

def tentukan_wilayah(nama_satker, jenis_satker):
    if jenis_satker == "Pusat": return "Pusat"
    s = nama_satker.lower()
    
    # Deteksi cepat kata kunci provinsi langsung
    if "jawa timur" in s or "jatim" in s: return "Jawa Timur"
    if "banten" in s: return "Banten"
    if "dki jakarta" in s or "jakarta" in s: return "DKI Jakarta"
    if "jawa barat" in s or "jabar" in s: return "Jawa Barat"
    if "jawa tengah" in s or "jateng" in s: return "Jawa Tengah"
    if "sumatera barat" in s or "sumbar" in s: return "Sumatera Barat"
    if "nusa tenggara timur" in s or "ntt" in s: return "Nusa Tenggara Timur"
    if "bengkulu" in s: return "Bengkulu"
    if "aceh" in s: return "Aceh"
    if "sumatera utara" in s or "sumut" in s: return "Sumatera Utara"
    if "riau" in s: return "Riau"
    if "kepulauan riau" in s or "kepri" in s: return "Kepulauan Riau"
    if "jambi" in s: return "Jambi"
    if "sumatera selatan" in s or "sumsel" in s: return "Sumatera Selatan"
    if "bangka belitung" in s or "babel" in s: return "Bangka Belitung"
    if "lampung" in s: return "Lampung"
    if "di yogyakarta" in s or "diy" in s or "jogja" in s: return "DI Yogyakarta"
    if "bali" in s: return "Bali"
    if "nusa tenggara barat" in s or "ntb" in s: return "Nusa Tenggara Barat"
    if "kalimantan barat" in s or "kalbar" in s: return "Kalimantan Barat"
    if "kalimantan tengah" in s or "kalteng" in s: return "Kalimantan Tengah"
    if "kalimantan selatan" in s or "kalsel" in s: return "Kalimantan Selatan"
    if "kalimantan timur" in s or "kaltim" in s: return "Kalimantan Timur"
    if "kalimantan utara" in s or "kaltara" in s: return "Kalimantan Utara"
    if "sulawesi utara" in s or "sulut" in s: return "Sulawesi Utara"
    if "gorontalo" in s: return "Gorontalo"
    if "sulawesi tengah" in s or "sulteng" in s: return "Sulawesi Tengah"
    if "sulawesi barat" in s or "sulbar" in s: return "Sulawesi Barat"
    if "sulawesi selatan" in s or "sulsel" in s: return "Sulawesi Selatan"
    if "sulawesi tenggara" in s or "sultra" in s: return "Sulawesi Tenggara"
    if "maluku utara" in s or "malut" in s: return "Maluku Utara"
    if "maluku" in s: return "Maluku"
    if "papua" in s: return "Papua"

    # KAMUS DETAIL KOTA/KABUPATEN SE-INDONESIA
    kamus_wilayah = {
        "Aceh": ["sabang", "lhokseumawe", "langsa", "subulussalam", "simeulue", "pidie", "bireuen", "gayo lues", "nagan raya", "bener meriah", "aceh"],
        "Sumatera Utara": ["medan", "binjai", "tebing tinggi", "pematang siantar", "tanjung balai", "sibolga", "padangsidimpuan", "nias", "mandailing natal", "tapanuli", "karo", "deli serdang", "langkat", "asahan", "labuhanbatu", "dairi", "toba", "samosir", "humbang hasundutan", "pakpak bharat", "simalungun", "batu bara", "padang lawas"],
        "Sumatera Barat": ["padang", "bukittinggi", "payakumbuh", "solok", "sawahlunto", "pariaman", "pasaman", "agam", "lima puluh kota", "tanah datar", "sijunjung", "dharmasraya", "pesisir selatan", "mentawai"],
        "Riau": ["pekanbaru", "dumai", "kampar", "rokan", "bengkalis", "siak", "pelalawan", "indragiri", "kuantan singingi", "meranti"],
        "Kepulauan Riau": ["batam", "tanjung pinang", "bintan", "karimun", "natuna", "anambas", "lingga"],
        "Jambi": ["sungai penuh", "kerinci", "merangin", "sarolangun", "batanghari", "muaro jambi", "tanjung jabung", "tebo", "bungo"],
        "Sumatera Selatan": ["palembang", "prabumulih", "lubuklinggau", "pagar alam", "banyuasin", "empat lawang", "lahat", "muara enim", "musi", "ogan", "penukal abab"],
        "Bangka Belitung": ["pangkal pinang", "bangka", "belitung"],
        "Bengkulu": ["rejang lebong", "mukomuko", "muko-muko", "kaur", "seluma", "kepahiang", "lebong"],
        "Lampung": ["metro", "tulang bawang", "tanggamus", "way kanan", "pesawaran", "pringsewu", "mesuji", "pesisir barat"],
        "Banten": ["tangerang", "serang", "cilegon", "pandeglang", "lebak"],
        "DKI Jakarta": ["kepulauan seribu"],
        "Jawa Barat": ["bandung", "bogor", "depok", "bekasi", "cimahi", "sukabumi", "cianjur", "garut", "tasikmalaya", "cirebon", "kuningan", "majalengka", "sumedang", "indramayu", "subang", "purwakarta", "karawang", "pangandaran", "banjar"],
        "Jawa Tengah": ["semarang", "surakarta", "solo", "salatiga", "tegal", "pekalongan", "banyumas", "cilacap", "purbalingga", "banjarnegara", "kebumen", "purworejo", "wonosobo", "boyolali", "klaten", "sukoharjo", "wonogiri", "karanganyar", "sragen", "grobogan", "blora", "rembang", "pati", "kudus", "jepara", "demak", "temanggung", "kendal", "batang", "pemalang", "brebes"],
        "DI Yogyakarta": ["sleman", "bantul", "gunungkidul", "kulon progo"],
        "Jawa Timur": ["sumenep", "bangkalan", "surabaya", "malang", "pacitan", "ponorogo", "situbondo", "lumajang", "blitar", "mojokerto", "sidoarjo", "gresik", "banyuwangi", "jember", "kediri", "tuban", "bojonegoro", "ngawi", "magetan", "madiun", "nganjuk", "trenggalek", "tulungagung", "jombang", "pasuruan", "probolinggo", "bondowoso", "lamongan", "pamekasan", "sampang", "batu"],
        "Bali": ["denpasar", "badung", "bangli", "buleleng", "gianyar", "jembrana", "karangasem", "klungkung", "tabanan"],
        "Nusa Tenggara Barat": ["mataram", "bima", "lombok", "sumbawa", "dompu"],
        "Nusa Tenggara Timur": ["kupang", "alor", "belu", "ende", "flores", "lembata", "manggarai", "ngada", "nagekeo", "rote", "sabu", "sikka", "sumba", "timor"],
        "Kalimantan Barat": ["pontianak", "singkawang", "sambas", "mempawah", "sanggau", "ketapang", "sintang", "kapuas hulu", "sekadau", "melawi", "kayong utara", "kubu raya"],
        "Kalimantan Tengah": ["palangka raya", "kotawaringin", "kapuas", "barito", "katingan", "seruyan", "sukamara", "lamandau", "gunung mas", "pulang pisau", "murung raya"],
        "Kalimantan Selatan": ["banjarmasin", "banjarbaru", "tanah laut", "kotabaru", "banjar", "tapin", "hulu sungai", "tabalong", "tanah bumbu", "balangan"],
        "Kalimantan Timur": ["samarinda", "balikpapan", "bontang", "paser", "kutai", "berau", "penajam", "mahakam"],
        "Kalimantan Utara": ["tarakan", "bulungan", "malinau", "nunukan", "tana tidung"],
        "Sulawesi Utara": ["manado", "bitung", "tomohon", "kotamobagu", "bolaang", "minahasa", "sangihe", "talaud", "sitaro"],
        "Gorontalo": ["boalemo", "bone bolango", "pohuwato"],
        "Sulawesi Tengah": ["palu", "banggai", "morowali", "poso", "donggala", "tolitoli", "toli-toli", "buol", "parigi", "tojo", "sigi"],
        "Sulawesi Barat": ["mamuju", "majene", "polewali mandar", "mamasa", "pasangkayu"],
        "Sulawesi Selatan": ["makassar", "parepare", "pare-pare", "palopo", "bantaeng", "barru", "bone", "bulukumba", "enrekang", "gowa", "jeneponto", "selayar", "luwu", "maros", "pangkajene", "pinrang", "sidenreng", "sinjai", "soppeng", "takalar", "tana toraja", "wajo"],
        "Sulawesi Tenggara": ["kendari", "baubau", "bau-bau", "buton", "muna", "konawe", "kolaka", "bombana", "wakatobi"],
        "Maluku": ["ambon", "tual", "buru", "seram", "aru", "tanimbar"],
        "Maluku Utara": ["ternate", "tidore", "halmahera", "sula", "morotai", "taliabu"],
        "Papua": ["jayapura", "biak", "yapen", "waropen", "sarmi", "keerom", "merauke", "boven digoel", "mappi", "asmat", "nabire", "mimika", "paniai", "dogiyai", "intan jaya", "deiyai", "manokwari", "sorong", "raja ampat", "fakfak", "fak-fak", "kaimana", "bintuni", "wondama", "tambrauw", "maybrat", "arfak", "jayawijaya", "bintang", "yahukimo", "tolikara", "mamberamo", "yalimo", "lanny jaya", "nduga", "puncak"]
    }
    
    for provinsi, daftar_kota in kamus_wilayah.items():
        for kota in daftar_kota:
            if re.search(fr'\b{re.escape(kota)}\b', s): return provinsi
                
    if "prov." in s: return s.split("prov.")[1].strip().title()
    if "provinsi " in s: return s.split("provinsi ")[1].strip().title()
    return "Wilayah Belum Diset"
