"""Liasse mails et factures (Capeos) : interface Streamlit.

On dépose des mails Outlook (.msg) ; l'app télécharge les factures et RIB donnés en lien,
tronque les factures Orange et les documents volumineux, et produit un seul PDF à imprimer.
Les fichiers sont traités en mémoire, le temps de la session, et ne sont jamais écrits sur disque.
"""
import hashlib
from datetime import datetime

import streamlit as st

import liasse_core as lc


def md(t):
    """Échappe le texte venant des mails pour l'affichage Markdown de Streamlit."""
    t = str(t)
    for c in "\\`*[]<>#|$~":
        t = t.replace(c, "\\" + c)
    return t

st.set_page_config(page_title="Liasse mails et factures", page_icon=":material/print:", layout="centered")

if "tour" not in st.session_state:
    st.session_state.tour = 0


def nouvelle_liasse():
    for k in ("cle", "res", "pdf_cle", "pdf"):
        st.session_state.pop(k, None)
    st.session_state.tour += 1


st.title("Liasse mails et factures")
st.caption("Les mails Outlook, leurs factures et RIB réunis dans un seul PDF à imprimer.")

fichiers = st.file_uploader(
    "Déposez les mails Outlook (.msg)",
    type=["msg"],
    accept_multiple_files=True,
    key=f"mails_{st.session_state.tour}",
    help="Dans Outlook, faites glisser les mails dans un dossier ou sur le bureau, puis déposez les fichiers ici.",
)

if not fichiers:
    st.markdown(
        f"""
**Ce que fait l'app, sans autre action de votre part :**
- chaque mail est imprimé sur sa propre page, suivi de ses pièces jointes (PDF, images), par date d'envoi ;
- la facture et le RIB donnés en lien dans le mail (« Facture : », « RIB : ») sont téléchargés automatiquement ; les justificatifs sont ignorés ;
- facture Orange : {lc.PAGES_DEBUT} premières pages, une page d'avis, puis la dernière page ;
- au-delà de {lc.SEUIL_PAGES} pages, la même troncature est proposée ;
- les mails en double sont écartés.
"""
    )
    st.caption("Les fichiers sont traités en mémoire le temps de la session, puis oubliés.")
    st.stop()

# ------------------------------------------------------------------ traitement (une fois par lot de mails)
empreinte = hashlib.sha256()
for f in sorted(fichiers, key=lambda x: x.name):
    empreinte.update(f.name.encode("utf-8"))
    empreinte.update(f.getvalue())
cle = empreinte.hexdigest()

if st.session_state.get("cle") != cle:
    with st.status("Traitement des mails…", expanded=False) as statut:
        def progression(i, n, m):
            statut.update(label=f"Mail {i} / {n} : {md(m.sujet[:70])}")

        res = lc.analyser([(f.name, f.getvalue()) for f in fichiers], progression)
        statut.update(label="Mails traités", state="complete")
    st.session_state.cle = cle
    st.session_state.res = res
    st.session_state.pop("pdf_cle", None)
res = st.session_state.res

if not res.a_imprimer:
    st.error("Aucun mail lisible : vérifiez qu'il s'agit bien de fichiers .msg enregistrés depuis Outlook.")
    for t in res.a_verifier:
        st.markdown(f"- {md(t)}")
    st.stop()

# ------------------------------------------------------------------ documents volumineux
tronquer = {}
if res.volumineux:
    st.subheader("Documents volumineux")
    st.caption(f"Au-delà de {lc.SEUIL_PAGES} pages : {lc.PAGES_DEBUT} premières pages, une page d'avis, puis la dernière page. Décochez pour tout imprimer.")
    for m, p in res.volumineux:
        tronquer[p.uid] = st.checkbox(
            f"Tronquer « {md(p.nom)} » ({p.pages} pages), mail du {lc.date_courte(m.date)}",
            value=True,
            key=f"t_{cle[:12]}_{p.uid}",
        )

# ------------------------------------------------------------------ PDF
pdf_cle = (cle, tuple(sorted(tronquer.items())))
if st.session_state.get("pdf_cle") != pdf_cle:
    with st.spinner("Assemblage du PDF…"):
        st.session_state.pdf = lc.assembler(res, tronquer)
    st.session_state.pdf_cle = pdf_cle
pdf, nb_pages, problemes = st.session_state.pdf

maintenant = datetime.now(lc.TZ) if lc.TZ else datetime.now()
st.download_button(
    f"Télécharger le PDF ({lc.pluriel(nb_pages, 'page')})",
    data=pdf,
    file_name=f"Liasse_mails_{maintenant:%Y-%m-%d_%Hh%M}.pdf",
    mime="application/pdf",
    type="primary",
    icon=":material/download:",
    use_container_width=True,
)

c1, c2, c3 = st.columns(3)
c1.metric("Mails imprimés", len(res.a_imprimer))
c2.metric("Doublons écartés", len(res.doublons))
c3.metric("À vérifier", len(res.a_verifier) + len(problemes))

a_verifier = res.a_verifier + [f"{t} n'a pas pu être ajouté, à imprimer à part" for t in problemes]
if a_verifier:
    st.warning("**À vérifier**\n\n" + "\n".join(f"- {md(t)}" for t in a_verifier))

if res.doublons:
    st.info("**Mails en double écartés**\n\n" + "\n".join(
        f"- {lc.date_courte(m.date)} · {md(m.sujet)} (identique au mail du {lc.date_courte(m.doublon_de.date)})" for m in res.doublons))

st.subheader("Détail par mail")
for i, m in enumerate(res.a_imprimer, 1):
    with st.expander(f"{i}. {lc.date_courte(m.date)} · {md(m.sujet)}"):
        st.caption(f"De : {md(m.de or '-')}")
        st.markdown("\n".join(f"- {md(ligne)}" for ligne in lc.resume_pieces(m)))

st.divider()
st.button("Nouvelle liasse", on_click=nouvelle_liasse, icon=":material/restart_alt:")
st.caption("Les fichiers sont traités en mémoire le temps de la session, puis oubliés. Téléchargements autorisés : " + ", ".join(lc.DOMAINES_AUTORISES) + ".")
