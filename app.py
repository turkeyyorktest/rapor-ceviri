import streamlit as st
import pandas as pd
from docx import Document
import io

# Sayfa ayarları
st.set_page_config(
    page_title="YorkTest Rapor Çevirici",
    page_icon="🇹🇷",
    layout="centered"
)

# ŞİFRE KORUMASI
def check_password():
    def password_entered():
        if st.session_state["password"] == "OxdXmX2vxM":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔐 Giriş")
        st.text_input(
            "Şifre", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.info("Lütfen şifrenizi girin")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔐 Giriş")
        st.text_input(
            "Şifre", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ Yanlış şifre!")
        return False
    else:
        return True

if not check_password():
    st.stop()

# Başlık
st.title("🇹🇷 YorkTest Rapor Çevirici")
st.markdown("**İngilizce DOCX raporlarını Türkçe'ye çevirin**")
st.markdown("---")

# Excel çeviri listesini yükle
@st.cache_data
def load_translations():
    df = pd.read_excel("Premium food&drink list_179 (1).xlsx")

    translation_dict = {}
    reverse_dict = {}

    for idx, row in df.iterrows():
        if idx == 0:
            continue
        english = str(row.iloc[0]).strip()
        turkish = str(row.iloc[1]).strip()

        if english and turkish and english != 'nan' and turkish != 'nan':
            translation_dict[english] = turkish
            reverse_dict[turkish] = english

            # Varyasyonlar
            translation_dict[english.lower()] = turkish
            for apos in ["'", "'", "`", "'"]:
                translation_dict[english.replace(apos, "'")] = turkish
                translation_dict[english.replace(apos, "")] = turkish

    return translation_dict, reverse_dict

try:
    translation_dict, reverse_dict = load_translations()
    sorted_foods = sorted(translation_dict.keys(), key=len, reverse=True)
    st.success(f"✅ {len(set(translation_dict.values()))} gıda çevirisi yüklendi!")
except Exception as e:
    st.error(f"❌ Çeviri listesi yüklenemedi: {e}")
    st.stop()

# Dosya yükleme
st.markdown("### 📤 1. DOCX Dosyasını Yükleyin")
uploaded_file = st.file_uploader(
    "İngilizce YorkTest raporunu seçin (DOCX formatında)",
    type=['docx'],
    help="Sadece .docx uzantılı dosyalar kabul edilir"
)

if uploaded_file is not None:
    st.success(f"✅ Dosya yüklendi: **{uploaded_file.name}**")

    # Çevir butonu
    st.markdown("### 🔄 2. Çeviriyi Başlatın")

    if st.button("🇹🇷 TÜRKÇE'YE ÇEVİR", type="primary", use_container_width=True):
        with st.spinner("⏳ Çeviri yapılıyor... Lütfen bekleyin..."):
            try:
                # DOCX'i aç
                doc = Document(io.BytesIO(uploaded_file.read()))

                translation_count = 0
                translated_foods = set()

                def translate_full_text(text):
                    if not text or len(text.strip()) < 2:
                        return text, 0

                    original = text
                    count = 0

                    for english_food in sorted_foods:
                        if english_food in text:
                            turkish_food = translation_dict[english_food]
                            if turkish_food not in text:
                                text = text.replace(english_food, turkish_food)
                                count += 1
                                translated_foods.add(f"{english_food} → {turkish_food}")

                    return text, count

                # Paragrafları çevir
                for para in doc.paragraphs:
                    full_para_text = para.text

                    if not full_para_text or len(full_para_text.strip()) < 2:
                        continue

                    new_para_text, count = translate_full_text(full_para_text)

                    if new_para_text != full_para_text and count > 0:
                        for run in para.runs:
                            run.text = ''
                        if para.runs:
                            para.runs[0].text = new_para_text
                        else:
                            para.add_run(new_para_text)
                        translation_count += count

                # Tabloları çevir
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            cell_text = cell.text

                            if not cell_text or len(cell_text.strip()) < 2:
                                continue

                            new_cell_text, count = translate_full_text(cell_text)

                            if new_cell_text != cell_text and count > 0:
                                if cell.paragraphs:
                                    para = cell.paragraphs[0]
                                    for run in para.runs:
                                        run.text = ''
                                    if para.runs:
                                        para.runs[0].text = new_cell_text
                                    else:
                                        para.add_run(new_cell_text)
                                translation_count += count

                # Dosyayı kaydet
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)

                # Başarı mesajı
                st.success("🎉 Çeviri tamamlandı!")
                st.info(f"📊 **{len(translated_foods)}** farklı gıda çevrildi")

                # İndirme butonu
                st.markdown("### 📥 3. Türkçe Dosyayı İndirin")

                # Orijinal dosya adından müşteri adını çıkar
                original_name = uploaded_file.name.replace('.docx', '')
                output_name = f"{original_name}_TURKCE.docx"

                st.download_button(
                    label="⬇️ TÜRKÇE DOCX'İ İNDİR",
                    data=output,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )

                # Çevrilen örnekler
                with st.expander("🔍 Çevrilen Gıdaları Görüntüle"):
                    for item in sorted(translated_foods)[:50]:
                        st.text(f"• {item}")
                    if len(translated_foods) > 50:
                        st.text(f"... ve {len(translated_foods) - 50} tane daha")

            except Exception as e:
                st.error(f"❌ Hata oluştu: {e}")
                st.error("Lütfen dosyanın doğru formatta olduğundan emin olun.")

else:
    st.info("👆 Lütfen yukarıdan bir DOCX dosyası yükleyin")

# Alt bilgi
st.markdown("---")
st.markdown("YorkTest Türkiye - Rapor Çeviri Sistemi", unsafe_allow_html=True)
