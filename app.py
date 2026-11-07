import base64
import html
from datetime import datetime
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Cm
from PIL import Image
from supabase import create_client

st.set_page_config(page_title="Quiz Soru Bankası")

# ===============================================================
# 1) AYARLAR (Supabase anahtarları secrets dosyasından okunur)
# ===============================================================
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
except Exception:
    st.error(
        "Supabase anahtarları bulunamadı. `.streamlit/secrets.toml` dosyasını "
        "(veya Streamlit Cloud'daki Secrets bölümünü) kontrol et."
    )
    st.stop()

# ===============================================================
# 2) HAFIZA (session_state): sadece geçici ekran bilgileri.
#    Sorular ve testler veritabanında durur, burada DEĞİL.
# ===============================================================
DEFAULTS = {
    "sb": None,             # bu kullanıcıya özel Supabase bağlantısı
    "user": None,           # giriş yapan kullanıcı
    "open_test": None,      # açık olan test: {"id": ..., "title": ...}
    "quiz_active": False,
    "quiz_questions": [],
    "quiz_index": 0,
    "quiz_score": 0,
    "quiz_id": 0,
    "form_id": 0,
    "notices": [],          # sağ üstte gösterilecek bildirimler
    "notice_n": 0,
    "edit_ver": {},         # soru başına sayaç: hatada alanları kayıtlı değere döndürür
    "open_q": None,         # hata sonrası açık kalacak soru
}
for _k, _v in DEFAULTS.items():
    st.session_state.setdefault(_k, _v)


def sb():
    """Her kullanıcı oturumu için AYRI bir bağlantı (başkasının verisine karışmasın)."""
    if st.session_state.sb is None:
        st.session_state.sb = create_client(SUPABASE_URL, SUPABASE_KEY)
    return st.session_state.sb


def logout(msg=None):
    try:
        sb().auth.sign_out()
    except Exception:
        pass
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    if msg:
        notify(msg, "info", 5)


def run(fn):
    """Veritabanı işlemini güvenle çalıştırır. Oturum süresi dolduysa yeniler."""
    try:
        return fn()
    except Exception as e:
        text = str(e).lower()
        if "jwt" in text or "expired" in text:
            try:
                sb().auth.refresh_session()
                return fn()
            except Exception:
                logout("Oturum süresi doldu, lütfen tekrar giriş yap.")
                st.rerun()
        notify(f"Veritabanı hatası: {e}", "error", 8)
        render_notices()
        st.stop()


# ===============================================================
# 3) VERİTABANI İŞLEMLERİ
# ===============================================================
def db_list_tests():
    return run(lambda: sb().table("tests")
               .select("id,title,created_at,questions(count)")
               .order("created_at", desc=True).execute().data)


def db_create_test(title):
    return run(lambda: sb().table("tests").insert({"title": title}).execute().data[0])


def db_rename_test(test_id, title):
    run(lambda: sb().table("tests").update({"title": title}).eq("id", test_id).execute())


def db_delete_test(test_id):
    run(lambda: sb().table("tests").delete().eq("id", test_id).execute())


def db_get_questions(test_id):
    return run(lambda: sb().table("questions").select("*")
               .eq("test_id", test_id).order("created_at").execute().data)


def db_add_question(test_id, text, choices, answer, image_b64):
    row = {"test_id": test_id, "text": text, "choices": choices,
           "answer": answer, "image_b64": image_b64}
    run(lambda: sb().table("questions").insert(row).execute())


def db_update_question(question_id, data):
    run(lambda: sb().table("questions").update(data).eq("id", question_id).execute())


def db_delete_question(question_id):
    run(lambda: sb().table("questions").delete().eq("id", question_id).execute())


# ===============================================================
# 4) YARDIMCI FONKSİYONLAR
# ===============================================================
def image_to_b64(raw_bytes):
    """Yüklenen resmi küçültüp metne çevirir (veritabanında saklamak için)."""
    img = Image.open(BytesIO(raw_bytes)).convert("RGB")
    img.thumbnail((800, 800))
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def b64_to_bytes(text):
    return base64.b64decode(text)


def validate_question(text, choices):
    if not text.strip():
        return "Soru metni boş bırakılamaz."
    if not all(choices):
        return "Dört şık da doldurulmalı."
    if len(set(choices)) < 4:
        return "Şıklar birbirinden farklı olmalı."
    return None


def make_word(title, questions, show_answers):
    doc = Document()
    doc.add_heading(title, 0)

    for i, q in enumerate(questions, 1):
        p = doc.add_paragraph()
        p.add_run(f"{i}. {q['text']}").bold = True

        if q.get("image_b64"):
            try:
                img = Image.open(BytesIO(b64_to_bytes(q["image_b64"]))).convert("RGB")
                w, h = img.size
                ratio = w / h
                if ratio >= 1:
                    width, height = Cm(5), Cm(5 / ratio)
                else:
                    width, height = Cm(5 * ratio), Cm(5)
                buf = BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                doc.add_picture(buf, width=width, height=height)
            except Exception:
                doc.add_paragraph("   [!] Resim eklenemedi.")

        for j, choice in enumerate(q["choices"]):
            doc.add_paragraph(f"   {chr(65 + j)}) {choice}")

        if show_answers:
            doc.add_paragraph().add_run(f"Doğru Cevap: {q['answer']}").italic = True

        doc.add_paragraph("-" * 30)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def notify(msg, kind=None, seconds=3):
    """Sağ üstte kısa süre görünüp kaybolan bildirim ekler."""
    if kind is None:
        kind = "delete" if "silindi" in msg else "success"
    st.session_state.setdefault("notices", [])
    st.session_state.notices.append((msg, kind, seconds))


def user_from(u):
    """Supabase kullanıcısından ekranda kullanılacak bilgileri çıkarır."""
    meta = getattr(u, "user_metadata", None) or {}
    return {"id": u.id, "email": u.email, "name": (meta.get("full_name") or "").strip()}


APP_CSS = (
    '[data-testid="stMainBlockContainer"],.block-container{padding-top:4.5rem!important}'
    ".banner{display:flex;align-items:center;gap:20px;flex-wrap:wrap;padding:12px 20px;"
    "border:1px solid rgba(128,128,128,.25);border-radius:12px;"
    "background:rgba(128,128,128,.07);margin-bottom:1.5rem}"
    ".banner-ad{font-size:1.05rem;font-weight:600;margin-right:12px}"
    ".banner-link{font-size:.9rem;color:color-mix(in srgb,currentColor 65%,transparent)}"
    ".selam{font-size:1.05rem;margin:0 0 .75rem}"
    'section[data-testid="stSidebar"] [data-testid="stButton"] button{width:100%}'
    '[class*="st-key-tbtn_"] button,[class*="st-key-tsel_"] button{justify-content:flex-start;'
    "text-align:left;background:transparent;border-color:transparent}"
    '[class*="st-key-tbtn_"] button p,[class*="st-key-tsel_"] button p{text-align:left}'
    '[class*="st-key-tbtn_"] button:hover{background:rgba(128,128,128,.12);border-color:transparent}'
    '[class*="st-key-tsel_"] button{background:rgba(128,128,128,.2);font-weight:600}'
    '[class*="st-key-save_"] button{border:1px solid #16a34a!important;background:transparent!important}'
    '[class*="st-key-save_"] button,[class*="st-key-save_"] button *{color:#16a34a!important}'
    '[class*="st-key-save_"] button:hover{background:rgba(22,163,74,.1)!important}'
    '[class*="st-key-del_"] button,.st-key-logout_box button{'
    "border:1px solid #dc2626!important;background:transparent!important}"
    '[class*="st-key-del_"] button,[class*="st-key-del_"] button *,'
    ".st-key-logout_box button,.st-key-logout_box button *{color:#dc2626!important}"
    '[class*="st-key-del_"] button:hover,.st-key-logout_box button:hover{background:rgba(220,38,38,.1)!important}'
    '[class*="st-key-del_"] button:disabled{opacity:.45}'
    '[data-testid="stSidebarUserContent"]>[data-testid="stVerticalBlock"]{min-height:calc(100vh - 8rem)}'
    '[data-testid="stSidebarUserContent"]>[data-testid="stVerticalBlock"]>:has(.st-key-logout_box),'
    ".st-key-logout_box{margin-top:auto}"
)


def render_notices():
    """Sayfa stillerini ve bekleyen bildirimleri basar (bildirim sağ üstte, süre dolunca kaybolur)."""
    notices = st.session_state.get("notices") or []
    st.session_state.notices = []
    css = APP_CSS
    items = ""
    if notices:
        n = st.session_state.get("notice_n", 0) + 1
        st.session_state.notice_n = n  # her seferinde yeni animasyon adı: animasyon baştan başlar
        colors = {"success": "#16a34a", "delete": "#475569", "error": "#dc2626",
                  "warning": "#d97706", "info": "#2563eb"}
        for msg, kind, secs in notices:
            color = colors.get(kind, colors["info"])
            items += (
                f'<div class="bildirim" style="background:{color};animation-duration:{secs}s">'
                f"{html.escape(str(msg))}</div>"
            )
        css += (
            ".bildirim-kutu{position:fixed;top:70px;right:16px;z-index:1000000;"
            "display:flex;flex-direction:column;gap:8px;"
            "max-width:min(380px,calc(100vw - 32px));pointer-events:none;}"
            ".bildirim{color:#fff;padding:12px 16px;border-radius:10px;font-size:15px;"
            "line-height:1.4;box-shadow:0 6px 18px rgba(0,0,0,.35);"
            f"animation-name:bildirim{n};animation-timing-function:ease;"
            "animation-fill-mode:forwards;}"
            f"@keyframes bildirim{n}"
            "{0%{opacity:0;transform:translateY(-12px);}"
            "8%{opacity:1;transform:none;}"
            "80%{opacity:1;}"
            "100%{opacity:0;visibility:hidden;}}"
        )
    st.markdown(
        f'<style>{css}</style><div class="bildirim-kutu">{items}</div>',
        unsafe_allow_html=True,
    )


def open_test(test):
    st.session_state.open_test = {"id": test["id"], "title": test["title"]}
    st.session_state.quiz_active = False
    st.session_state.form_id += 1
    st.rerun()


# ===============================================================
# 5) GİRİŞ / KAYIT EKRANI
# ===============================================================
def show_login():
    st.title("Quiz Soru Bankası", anchor=False)
    tab_in, tab_up = st.tabs(["Giriş yap", "Kayıt ol"])

    with tab_in:
        with st.form("giris_formu"):
            email = st.text_input("E-posta")
            password = st.text_input("Şifre", type="password")
            go = st.form_submit_button("Giriş yap")
        if go:
            ok = False
            try:
                res = sb().auth.sign_in_with_password(
                    {"email": email.strip(), "password": password}
                )
                st.session_state.user = user_from(res.user)
                ok = True
            except Exception:
                notify("E-posta veya şifre hatalı (ya da e-postanı henüz onaylamadın).", "error", 5)
            if ok:
                st.rerun()

    with tab_up:
        with st.form("kayit_formu"):
            name2 = st.text_input("İsim", key="kayit_isim")
            email2 = st.text_input("E-posta", key="kayit_email")
            pw1 = st.text_input("Şifre (en az 6 karakter)", type="password", key="kayit_pw1")
            pw2 = st.text_input("Şifre (tekrar)", type="password", key="kayit_pw2")
            go2 = st.form_submit_button("Kayıt ol")
        if go2:
            ok = False
            if not name2.strip():
                notify("İsim boş bırakılamaz.", "error")
            elif not email2.strip():
                notify("E-posta boş bırakılamaz.", "error")
            elif len(pw1) < 6:
                notify("Şifre en az 6 karakter olmalı.", "error")
            elif pw1 != pw2:
                notify("Şifreler aynı değil.", "error")
            else:
                try:
                    res = sb().auth.sign_up({
                        "email": email2.strip(),
                        "password": pw1,
                        "options": {"data": {"full_name": name2.strip()[:40]}},
                    })
                    if res.session is None:
                        notify(
                            "Hesap oluşturuldu. E-postana gelen onay bağlantısına "
                            "tıkla, sonra giriş yap.", "success", 8
                        )
                    else:
                        st.session_state.user = user_from(res.user)
                        ok = True
                except Exception as e:
                    notify(f"Kayıt yapılamadı: {e}", "error", 6)
            if ok:
                st.rerun()


# ===============================================================
# 6) TESTLERİM SAYFASI
# ===============================================================
def page_tests():
    st.title("Test oluştur", anchor=False)
    st.caption("Yeni bir test oluştur. Testlerin sol menüde listelenir.")

    with st.form("yeni_test_formu", clear_on_submit=True):
        title = st.text_input("Test adı")
        create = st.form_submit_button("Test oluştur")
    if create:
        if not title.strip():
            notify("Test adı boş bırakılamaz.", "error")
        else:
            test = db_create_test(title.strip())
            notify(f"Test oluşturuldu: {title.strip()}")
            open_test(test)


# ===============================================================
# 7) TEK BİR TESTİN SAYFASI
# ===============================================================
def page_test():
    test = st.session_state.open_test

    st.title(test["title"], anchor=False)
    questions = db_get_questions(test["id"])
    st.caption(f"{len(questions)} soru · yaptığın her değişiklik otomatik kaydedilir")

    tab_q, tab_play, tab_word, tab_set = st.tabs(
        ["Sorular", "Testi Çöz", "Word'e Aktar", "Ayarlar"]
    )

    # ------------------------- SORULAR -------------------------
    with tab_q:
        st.subheader("Yeni soru ekle", anchor=False)
        fid = st.session_state.form_id
        with st.form(f"soru_formu_{fid}"):
            text = st.text_area("Soru metni", key=f"text_{fid}")
            a = st.text_input("A şıkkı", key=f"a_{fid}")
            b = st.text_input("B şıkkı", key=f"b_{fid}")
            c = st.text_input("C şıkkı", key=f"c_{fid}")
            d = st.text_input("D şıkkı", key=f"d_{fid}")
            correct = st.selectbox("Doğru cevap", list("ABCD"), key=f"correct_{fid}")
            image_file = st.file_uploader(
                "Resim (isteğe bağlı)", type=["png", "jpg", "jpeg"], key=f"img_{fid}"
            )
            submitted = st.form_submit_button("Soruyu Ekle")

        if submitted:
            choices = [a.strip(), b.strip(), c.strip(), d.strip()]
            error = validate_question(text, choices)
            img_b64 = None
            if not error and image_file is not None:
                try:
                    img_b64 = image_to_b64(image_file.getvalue())
                except Exception:
                    error = "Yüklenen dosya geçerli bir resim değil."
            if error:
                notify(error, "error")
            else:
                db_add_question(test["id"], text.strip(), choices,
                                choices["ABCD".index(correct)], img_b64)
                st.session_state.form_id += 1
                notify("Soru kaydedildi.")
                st.rerun()

        st.divider()
        st.subheader("Mevcut sorular", anchor=False)
        if not questions:
            st.info("Bu testte henüz soru yok.")

        for i, q in enumerate(questions, 1):
            qid = q["id"]
            ver = st.session_state.edit_ver.get(qid, 0)
            sfx = f"{qid}_{ver}"  # sürüm değişince alanlar kayıtlı değerle yeniden oluşur
            with st.expander(f"{i}. {q['text'][:60]}",
                             expanded=(st.session_state.open_q == qid)):
                with st.form(f"duzenle_{sfx}"):
                    e_text = st.text_area("Soru metni", value=q["text"], key=f"et_{sfx}")
                    e_choices = [
                        st.text_input(f"{'ABCD'[j]} şıkkı", value=q["choices"][j],
                                      key=f"ec{j}_{sfx}")
                        for j in range(4)
                    ]
                    cur = q["choices"].index(q["answer"]) if q["answer"] in q["choices"] else 0
                    e_letter = st.selectbox("Doğru cevap", list("ABCD"), index=cur,
                                            key=f"ea_{sfx}")
                    remove_img = False
                    if q.get("image_b64"):
                        st.image(b64_to_bytes(q["image_b64"]), width=200)
                        remove_img = st.checkbox("Resmi kaldır", key=f"er_{sfx}")
                    new_img = st.file_uploader("Resim ekle / değiştir",
                                               type=["png", "jpg", "jpeg"], key=f"ei_{sfx}")
                    col1, col2 = st.columns(2)
                    with col1:
                        with st.container(key=f"save_{qid}"):
                            save = st.form_submit_button("Kaydet")
                    with col2:
                        with st.container(key=f"del_{qid}"):
                            delete = st.form_submit_button("Sil")

                if save:
                    choices = [x.strip() for x in e_choices]
                    error = validate_question(e_text, choices)
                    update = {}
                    if not error:
                        update = {"text": e_text.strip(), "choices": choices,
                                  "answer": choices["ABCD".index(e_letter)]}
                        if new_img is not None:
                            try:
                                update["image_b64"] = image_to_b64(new_img.getvalue())
                            except Exception:
                                error = "Yüklenen dosya geçerli bir resim değil."
                        elif remove_img:
                            update["image_b64"] = None
                    if error:
                        st.session_state.edit_ver[qid] = ver + 1
                        st.session_state.open_q = qid
                        notify(error + " Değişiklik kaydedilmedi, kayıtlı hâli geri yüklendi.",
                               "error", 5)
                        st.rerun()
                    else:
                        db_update_question(qid, update)
                        st.session_state.open_q = None
                        notify("Değişiklik kaydedildi.")
                        st.rerun()

                if delete:
                    db_delete_question(qid)
                    st.session_state.open_q = None
                    notify("Soru silindi.")
                    st.rerun()

    # ------------------------- TEST ÇÖZ -------------------------
    with tab_play:
        if not questions and not st.session_state.quiz_active:
            st.info("Önce soru eklemelisin.")

        elif not st.session_state.quiz_active:
            st.write(f"Toplam **{len(questions)}** soru ile test başlayacak.")
            if st.button("Testi Başlat"):
                st.session_state.quiz_active = True
                st.session_state.quiz_questions = list(questions)
                st.session_state.quiz_index = 0
                st.session_state.quiz_score = 0
                st.session_state.quiz_id += 1
                st.rerun()

        else:
            quiz_qs = st.session_state.quiz_questions
            idx = st.session_state.quiz_index
            total = len(quiz_qs)

            if idx >= total:
                score = st.session_state.quiz_score
                st.subheader("Test bitti", anchor=False)
                st.metric("Skorun", f"{int(score / total * 100)}")
                st.write(f"Toplam {total} soruda {score} doğru yaptın.")
                if st.button("Bitir"):
                    st.session_state.quiz_active = False
                    st.rerun()
            else:
                q = quiz_qs[idx]
                st.progress(idx / total, text=f"Soru {idx + 1} / {total}")
                st.subheader(q["text"], anchor=False)
                if q.get("image_b64"):
                    st.image(b64_to_bytes(q["image_b64"]), width=300)

                choice = st.radio("Cevabın:", q["choices"], index=None,
                                  key=f"cevap_{st.session_state.quiz_id}_{idx}")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Cevabı Gönder"):
                        if choice is None:
                            notify("Lütfen bir seçenek seç.", "warning")
                        else:
                            if choice == q["answer"]:
                                st.session_state.quiz_score += 1
                            st.session_state.quiz_index += 1
                            st.rerun()
                with col2:
                    if st.button("Testi Bırak"):
                        st.session_state.quiz_active = False
                        st.rerun()

    # ------------------------- WORD -------------------------
    with tab_word:
        if not questions:
            st.info("Aktarılacak soru yok.")
        else:
            name_input = st.text_input("Dosya adı", value=test["title"])
            show_answers = st.checkbox("Doğru cevaplar Word dosyasına yazılsın")
            clean = "".join(ch for ch in name_input if ch.isalnum() or ch in " _-").strip()
            if not clean:
                clean = "Quiz_Sorulari"
            stamp = datetime.now().strftime("%d-%m-%Y_%H.%M.%S")
            st.download_button(
                "Word dosyasını indir",
                data=make_word(test["title"], questions, show_answers),
                file_name=f"{clean}_{stamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    # ------------------------- AYARLAR -------------------------
    with tab_set:
        st.subheader("Test adını değiştir", anchor=False)
        new_title = st.text_input("Test adı", value=test["title"], key="rename_input")
        with st.container(key="save_rename"):
            rename_clicked = st.button("Adı kaydet")
        if rename_clicked:
            if not new_title.strip():
                notify("Test adı boş bırakılamaz.", "error")
            else:
                db_rename_test(test["id"], new_title.strip())
                st.session_state.open_test["title"] = new_title.strip()
                notify("Test adı güncellendi.")
                st.rerun()

        st.divider()
        st.subheader("Testi sil", anchor=False)
        st.caption("Testle birlikte içindeki tüm sorular silinir. Bu geri alınamaz.")
        confirm = st.checkbox("Bu testi silmek istediğimden eminim.")
        with st.container(key="del_test"):
            delete_clicked = st.button("Testi Sil", disabled=not confirm)
        if delete_clicked:
            db_delete_test(test["id"])
            st.session_state.open_test = None
            st.session_state.quiz_active = False
            notify("Test silindi.")
            st.rerun()


def render_banner():
    st.markdown(
        '<div class="banner"><span class="banner-ad">Şirket adı</span>'
        '<span class="banner-link">Testlerim</span>'
        '<span class="banner-link">Hakkımızda</span>'
        '<span class="banner-link">İletişim</span>'
        '<span class="banner-link">SSS</span></div>',
        unsafe_allow_html=True,
    )


def render_sidebar():
    user = st.session_state.user
    name = user.get("name") or user["email"].split("@")[0]
    with st.sidebar:
        st.markdown(
            f'<p class="selam">Merhaba, <b>{html.escape(name)}</b></p>',
            unsafe_allow_html=True,
        )

        if not user.get("name"):
            with st.expander("Adını ekle"):
                new_name = st.text_input("İsmin", key="isim_ekle")
                with st.container(key="save_isim"):
                    add_clicked = st.button("Kaydet", key="isim_kaydet")
                if add_clicked:
                    saved = False
                    if not new_name.strip():
                        notify("İsim boş bırakılamaz.", "error")
                    else:
                        try:
                            sb().auth.update_user({"data": {"full_name": new_name.strip()[:40]}})
                            st.session_state.user["name"] = new_name.strip()[:40]
                            notify("İsmin kaydedildi.")
                            saved = True
                        except Exception as e:
                            notify(f"İsim kaydedilemedi: {e}", "error", 6)
                    if saved:
                        st.rerun()

        st.caption("Testlerim")
        if st.button("Yeni test oluştur", key="yeni_test_btn"):
            st.session_state.open_test = None
            st.session_state.quiz_active = False
            st.rerun()

        tests = db_list_tests()
        current = st.session_state.open_test
        if not tests:
            st.caption("Henüz testin yok.")
        for t in tests:
            try:
                count = t["questions"][0]["count"]
            except Exception:
                count = 0
            selected = bool(current) and current["id"] == t["id"]
            with st.container(key=f"{'tsel' if selected else 'tbtn'}_{t['id']}"):
                if st.button(f"{t['title']} ({count})", key=f"ac_{t['id']}"):
                    open_test(t)

        with st.container(key="logout_box"):
            if st.button("Çıkış yap", key="logout_btn"):
                logout()
                st.rerun()


# ===============================================================
# 8) ANA AKIŞ
# ===============================================================
render_banner()

if st.session_state.user is None:
    show_login()
else:
    render_sidebar()
    if st.session_state.open_test is None:
        page_tests()
    else:
        page_test()

# Stiller ve bekleyen bildirimler en sonda basılır
render_notices()