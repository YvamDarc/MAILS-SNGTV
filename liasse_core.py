"""
Liasse mails et factures (Capeos) : moteur de traitement, sans interface.

Réunit dans un seul PDF, classés par date d'envoi :
  - chaque mail Outlook (.msg) imprimé sur sa propre page,
  - ses pièces jointes PDF et images,
  - la facture et le RIB donnés en lien dans le mail (« Facture : », « RIB : »), téléchargés
    automatiquement depuis les sites autorisés ; les liens « Justificatif » sont ignorés.

Troncature : facture Orange = 4 premières pages + page d'avis + dernière page.
Au-delà de SEUIL_PAGES pages, la même troncature est proposée.
Les mails en double sont écartés ; une pièce identique dans deux mails est imprimée et signalée.

Confidentialité : tout est traité en mémoire, rien n'est écrit sur disque ni journalisé.
"""

# ======================= Réglages =======================
SEUIL_PAGES = 10                         # au-delà, proposer de tronquer (hors Orange)
PAGES_DEBUT = 4                          # pages gardées au début d'un document tronqué
DOMAINES_AUTORISES = ("sngtv.org",)      # seuls sites d'où les factures et RIB sont téléchargés
TAILLE_MAX_TELECHARGEMENT = 60 * 1024 * 1024
# ========================================================

import base64
import hashlib
import io
import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlsplit, unquote


import logging
logging.disable(logging.WARNING)  # extract_msg et pypdf sont bavards

import extract_msg
import requests
from fpdf import FPDF
from pypdf import PdfReader, PdfWriter

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Europe/Paris")
except Exception:
    TZ = None

ICI = Path(__file__).resolve().parent
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]


# ----------------------------------------------------------------- affichage
def pluriel(n, mot, pl=None):
    return f"{n} {mot if n <= 1 else (pl or mot + 's')}"


def date_locale(d):
    if d is None:
        return None
    if d.tzinfo is None:
        return d
    return d.astimezone(TZ) if TZ else d


def date_courte(d):
    d = date_locale(d)
    return d.strftime("%d/%m/%Y à %H:%M") if d else "date inconnue"


def date_longue(d):
    d = date_locale(d)
    if not d:
        return "date inconnue"
    return f"{JOURS[d.weekday()]} {d.day} {MOIS[d.month - 1]} {d.year} à {d:%H:%M}"


def empreinte(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ----------------------------------------------------------------- liens
URL_RE = re.compile(r'https?://[^\s<>"]+')


def _b64url(s):
    s = s.replace("-", "+").replace("_", "/")
    s += "=" * (-len(s) % 4)
    return base64.b64decode(s).decode("utf-8", "replace")


def deballer_lien(url):
    """Retire les protections Outlook Safe Links et Mailinblack pour retrouver l'adresse d'origine."""
    u = url
    for _ in range(6):
        try:
            p = urlsplit(u)
        except ValueError:
            return None
        hote = (p.hostname or "").lower()
        q = parse_qs(p.query)
        if hote.endswith("safelinks.protection.outlook.com") and q.get("url"):
            u = q["url"][0]
            continue
        if "mailinblack.com" in hote and q.get("key"):
            try:
                cible = json.loads(_b64url(q["key"][0])).get("url")
            except Exception:
                return None
            if not cible:
                return None
            u = cible
            continue
        return u
    return u


def hote_de(url):
    try:
        h = (urlsplit(url).hostname or "").lower()
    except ValueError:
        return ""
    return h[4:] if h.startswith("www.") else h


def fichier_de(url):
    try:
        seg = [s for s in urlsplit(url).path.split("/") if s]
    except ValueError:
        return ""
    if seg and re.search(r"\.[a-z0-9]{2,5}$", seg[-1], re.I):
        return unquote(seg[-1])
    return ""


def type_de_lien(texte):
    s = (texte or "")[-48:].lower()
    if "justificatif" in s:
        return None
    if re.search(r"\brib\b|relev[ée] d.identit[ée] bancaire", s):
        return "RIB"
    if "factur" in s:
        return "Facture"
    return None


def liens_facture_rib(corps):
    lignes = corps.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    trouves, vus = [], set()
    for i, ligne in enumerate(lignes):
        for m in URL_RE.finditer(ligne):
            avant = re.sub(r"[\s<]+$", "", re.sub(r"^[\s>]+", "", ligne[: m.start()]))
            if not avant:
                for j in range(i - 1, max(-1, i - 3), -1):
                    t = re.sub(r"^[\s>]+", "", lignes[j]).strip()
                    if t:
                        if len(t) <= 40:
                            avant = t
                        break
            etiquette = type_de_lien(avant)
            if not etiquette:
                continue
            url = re.sub(r"[>.,;)\]]+$", "", m.group(0))
            cible = deballer_lien(url)
            hote = hote_de(cible) if cible else ""
            if not cible or "." not in hote:
                continue  # lien coupé par la citation d'un ancien mail
            cle = (etiquette, cible)
            if cle in vus:
                continue
            vus.add(cle)
            trouves.append({"type": etiquette, "url": url, "cible": cible, "hote": hote, "fichier": fichier_de(cible)})
    trouves.sort(key=lambda l: 0 if l["type"] == "Facture" else 1)
    return trouves


def nb_justificatifs(corps):
    return len(re.findall(r"^[ \t]*justificatifs?[ \t]*:", corps, re.I | re.M))


# ----------------------------------------------------------------- corps du mail
CAR_URL = re.compile(r"^[A-Za-z0-9%._~:/?#\[\]@!$&'()*+,;=-]+$")


def nettoyer_corps(corps):
    sortie, suite_lien = [], False
    for brut in corps.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        m = re.match(r"^(\s*>)+\s?", brut)
        q = m.group(0) if m else ""
        c = re.sub(r"[\u00a0\u200b\u200c\u200d\ufeff]", " ", brut[len(q):])
        t = c.strip()
        if suite_lien and t and not re.search(r"\s", t) and CAR_URL.match(t) and (re.search(r"[%=&]", t) or len(t) >= 30):
            continue  # morceau d'un long lien coupé
        suite_lien = bool(re.search(r"https?://\S+\s*>?\s*$", c))
        if re.search(r"ATTENTION\s*:\s*Cet e-mail provient d.une personne externe", c, re.I):
            continue

        def remplace(mo):
            cible = deballer_lien(mo.group(1))
            h = hote_de(cible) if cible else ""
            return "[lien" + (f" {h}" if "." in h else "") + "]"

        c = re.sub(r"\s*<mailto:[^>]*>", "", c, flags=re.I)
        c = re.sub(r'<?(https?://[^\s<>"]+)>?', remplace, c)
        c = re.sub(r"(\[lien[^\]]*\])(\s*\1)+", r"\1", c).rstrip()
        if re.fullmatch(r"\s*(\[lien[^\]]*\]\s*)+", c):
            if sortie and sortie[-1].rstrip().endswith(":"):
                sortie[-1] = sortie[-1].rstrip() + " " + c.strip()
            continue
        sortie.append((q.strip() + " " + c).rstrip() if q.strip() else c)
    res, vide = [], 0
    for l in sortie:
        if not re.sub(r"^[>\s]+", "", l).strip() or l.strip() == "'":
            vide += 1
            if vide > 1:
                continue
            res.append("")
        else:
            vide = 0
            res.append(l)
    while res and not res[0]:
        res.pop(0)
    while res and not res[-1]:
        res.pop()
    return "\n".join(res)


# ----------------------------------------------------------------- pièces
def genre(data, ext):
    if data[:4] == b"%PDF" or (ext == "pdf" and b"%PDF" in data[:1024]):
        return "pdf"
    if data[:8] == b"\x89PNG\r\n\x1a\n" or data[:3] == b"\xff\xd8\xff" or data[:4] == b"GIF8" or (data[:4] == b"RIFF" and data[8:12] == b"WEBP"):
        return "image"
    if data[:2] == b"BM" and ext == "bmp":
        return "image"
    if ext in ("htm", "html") and len(data) < 4096:
        return "technique"
    return "autre"


class Piece:
    def __init__(self, nom, data, origine="pj", dans="", cachee=False):
        self.nom = nom or "pièce jointe"
        self.data = data
        self.ext = (Path(self.nom).suffix[1:] or "").lower()
        self.genre = genre(data, self.ext)
        self.origine = origine          # pj, lien
        self.dans = dans                # sujet du mail joint qui la contenait
        self.empreinte = empreinte(data)
        self.pages = 0
        self.orange = False
        self.tronquer = False
        self.motif = ""
        self.imprimer = self.genre in ("pdf", "image")
        self.raison = ""
        self.doublon_de = None
        self.lecteur = None
        if self.genre == "autre":
            self.raison = "format non imprimable, à imprimer à part"
        if self.genre == "technique":
            self.raison = "fichier technique Outlook"
        if self.genre == "image" and (cachee or (re.fullmatch(r"image\d{3}\.(png|jpe?g|gif)", self.nom, re.I) and len(data) < 60000) or len(data) < 8000):
            self.imprimer = False
            self.raison = "image de signature probable"

    def analyser(self, mail):
        if self.genre == "image":
            self.pages = 1
            return
        if self.genre != "pdf":
            return
        try:
            r = PdfReader(io.BytesIO(self.data), strict=False)
            if r.is_encrypted and not r.decrypt(""):
                raise PermissionError
            self.pages = len(r.pages)
            self.lecteur = r
        except PermissionError:
            self.imprimer, self.raison = False, "PDF protégé par mot de passe, à imprimer à part"
            return
        except Exception:
            self.imprimer, self.raison = False, "PDF illisible, à imprimer à part"
            return
        texte = ""
        for p in r.pages[:2]:
            try:
                texte += (p.extract_text() or "") + "\n"
            except Exception:
                pass
        self.orange = est_orange(self, mail, texte)


def est_orange(piece, mail, texte):
    if "orange" in piece.nom.lower():
        return True
    if re.search(r"@([a-z0-9-]+\.)*orange(-business)?\.(fr|com)$", mail.de_email or "", re.I):
        return True
    t = re.sub(r"\b\d{5}\s+orange\b", " ", texte, flags=re.I)
    return bool(re.search(r"\borange\b", t, re.I)) and bool(re.search(r"factur", texte, re.I))


# ----------------------------------------------------------------- mails
class Mail:
    def __init__(self, chemin):
        self.chemin = Path(chemin)
        self.sujet = ""
        self.de = ""
        self.de_email = ""
        self.a = ""
        self.cc = ""
        self.date = None
        self.corps = ""
        self.propre = ""
        self.pieces = []
        self.liens = []          # dicts + "piece" + "statut"
        self.justificatifs = 0
        self.doublon_de = None
        self.cle = ""
        self.erreur = ""

    def a_des_pieces(self):
        return any(p.genre in ("pdf", "image") and p.raison != "image de signature probable" for p in self.pieces)

    def imprimees(self):
        out = [p for p in self.pieces if p.imprimer]
        out += [l["piece"] for l in self.liens if l.get("piece") is not None and l["piece"].imprimer]
        return out


def _texte(v):
    if v is None:
        return ""
    if isinstance(v, bytes):
        return v.decode("utf-8", "replace")
    return str(v)


def lire_msg(source, mail, profondeur=0, dans=""):
    msg = extract_msg.openMsg(bytes(source)) if isinstance(source, (bytes, bytearray)) else source
    corps = _texte(getattr(msg, "body", "") or "")
    if not corps.strip():
        try:
            html = _texte(msg.htmlBody)
            html = re.sub(r"(?is)<(script|style).*?</\1>", "", html)
            html = re.sub(r'(?is)<a[^>]+href="(https?:[^"]+)"[^>]*>(.*?)</a>', r"\2 <\1>", html)
            html = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", html)
            from html import unescape
            corps = unescape(re.sub(r"<[^>]+>", "", html))
        except Exception:
            pass
    autres_corps = []
    for att in getattr(msg, "attachments", []) or []:
        try:
            data = att.data
            if hasattr(data, "attachments"):  # mail joint
                if profondeur < 3:
                    sous = lire_msg(data, mail, profondeur + 1, _texte(getattr(data, "subject", "")) or "mail joint")
                    autres_corps.append(sous)
                continue
            if not isinstance(data, (bytes, bytearray)):
                continue
            nom = _texte(getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or getattr(att, "name", None) or "pièce jointe")
            mail.pieces.append(Piece(nom, bytes(data), "pj", dans, bool(getattr(att, "hidden", False))))
        except Exception:
            mail.pieces.append(Piece(_texte(getattr(att, "longFilename", "")) or "pièce jointe", b"", "pj", dans))
            mail.pieces[-1].imprimer, mail.pieces[-1].raison = False, "pièce illisible"
    if profondeur == 0:
        mail.sujet = _texte(msg.subject).strip() or "(sans objet)"
        mail.de = _texte(msg.sender).strip()
        m = re.search(r"<([^>]+@[^>]+)>", mail.de) or re.search(r"[\w.+-]+@[\w.-]+", mail.de)
        mail.de_email = (m.group(1) if m and m.lastindex else (m.group(0) if m else "")).strip()
        mail.a = _texte(msg.to).strip()
        mail.cc = _texte(msg.cc).strip()
        d = getattr(msg, "date", None)
        if isinstance(d, str):
            try:
                from email.utils import parsedate_to_datetime
                d = parsedate_to_datetime(d)
            except Exception:
                d = None
        mail.date = d if isinstance(d, datetime) else None
        try:
            msg.close()
        except Exception:
            pass
    return "\n".join([corps] + autres_corps)


def charger(nom, data):
    mail = Mail(nom)
    try:
        corps = lire_msg(data, mail)
    except Exception:
        mail.erreur = "fichier illisible (est-ce bien un mail Outlook .msg ?)"
        mail.sujet = nom
        return mail
    mail.corps = corps
    mail.propre = nettoyer_corps(corps)
    mail.liens = liens_facture_rib(corps)
    mail.justificatifs = nb_justificatifs(corps)
    sujet = re.sub(r"^(\s*(\[[^\]]*\]|tr|re|fwd?|réf)\s*:?\s*)+", "", mail.sujet, flags=re.I).lower()
    base = sujet + "|" + re.sub(r"\s+", " ", mail.propre) + "|" + ",".join(sorted(p.empreinte for p in mail.pieces))
    mail.cle = empreinte(base.encode("utf-8"))
    return mail


# ----------------------------------------------------------------- téléchargement
def domaine_autorise(url):
    h = hote_de(url)
    return any(h == d or h.endswith("." + d) for d in DOMAINES_AUTORISES)


SESSION = requests.Session()
SESSION.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) LiasseMails/1.0"


def telecharger(url):
    """Renvoie (octets, None) ou (None, raison). N'accepte que des PDF ou des images."""
    if not url.lower().startswith(("https://", "http://")):
        return None, "adresse non valide"
    if not domaine_autorise(url):
        return None, "site non autorisé (" + (hote_de(url) or "?") + "), à télécharger à la main"
    try:
        with SESSION.get(url, timeout=(15, 60), stream=True, allow_redirects=True) as r:
            if not domaine_autorise(r.url):
                return None, "redirection vers un site non autorisé"
            if r.status_code != 200:
                return None, f"le site a répondu {r.status_code}"
            data = bytearray()
            for bloc in r.iter_content(65536):
                data.extend(bloc)
                if len(data) > TAILLE_MAX_TELECHARGEMENT:
                    return None, "fichier trop volumineux"
    except requests.exceptions.SSLError:
        return None, "erreur de certificat"
    except requests.exceptions.ConnectionError:
        return None, "site injoignable (connexion ou proxy)"
    except requests.exceptions.Timeout:
        return None, "délai dépassé"
    except Exception as e:
        return None, f"erreur ({type(e).__name__})"
    data = bytes(data)
    if genre(data, "pdf") == "pdf" or genre(data, "") == "image":
        return data, None
    return None, "le lien n'a pas renvoyé de PDF (page web ou connexion requise)"


# ----------------------------------------------------------------- PDF
def police():
    regs = [(ICI / "polices" / "Carlito-Regular.ttf", ICI / "polices" / "Carlito-Bold.ttf"),
            (Path(r"C:\Windows\Fonts\calibri.ttf"), Path(r"C:\Windows\Fonts\calibrib.ttf")),
            (Path(r"C:\Windows\Fonts\arial.ttf"), Path(r"C:\Windows\Fonts\arialbd.ttf"))]
    for r, b in regs:
        if r.exists() and b.exists():
            return r, b
    return None, None


POLICE_R, POLICE_B = police()
BLEU = (0, 61, 114)
VERT = (132, 171, 157)
FRAMBOISE = (211, 52, 84)
ENCRE = (16, 34, 54)
GRIS = (95, 108, 122)
FILET = (214, 222, 230)


def propre_pdf(t):
    t = unicodedata.normalize("NFC", str(t or ""))
    t = re.sub(r"[\u200b-\u200f\u2060\ufeff\u00ad\r]", "", t).replace("\t", "    ")
    t = "".join(ch for ch in t if ord(ch) < 0x10000 and (ch == "\n" or unicodedata.category(ch)[0] != "C"))
    if POLICE_R is None:
        t = t.translate(str.maketrans({"’": "'", "‘": "'", "“": '"', "”": '"', "–": "-", "—": "-", "…": "...", "€": "EUR", "œ": "oe", "Œ": "OE"}))
        t = t.encode("latin-1", "replace").decode("latin-1")
    return t


class Page(FPDF):
    def __init__(self, entete="", date_txt="", pied=""):
        super().__init__(format="A4", unit="mm")
        self.entete, self.date_txt, self.pied = entete, date_txt, pied
        self.set_margins(19, 24, 19)
        self.set_auto_page_break(True, 18)
        if POLICE_R:
            self.add_font("L", "", str(POLICE_R))
            self.add_font("L", "B", str(POLICE_B))
            self.f = "L"
        else:
            self.f = "helvetica"

    def header(self):
        if not self.entete:
            return
        self.set_y(11)
        self.set_font(self.f, "B", 8)
        self.set_text_color(*BLEU)
        self.cell(0, 4, propre_pdf(self.entete))
        self.set_font(self.f, "", 8)
        self.set_text_color(*GRIS)
        self.set_x(self.l_margin)
        self.cell(0, 4, propre_pdf(self.date_txt), align="R")
        self.set_draw_color(*VERT)
        self.set_line_width(0.45)
        self.line(self.l_margin, 16.5, self.w - self.r_margin, 16.5)
        self.set_y(23)

    def footer(self):
        if not self.pied:
            return
        self.set_y(-12)
        self.set_font(self.f, "", 7.5)
        self.set_text_color(*GRIS)
        pied = propre_pdf(self.pied)
        while len(pied) > 4 and self.get_string_width(pied) > self.epw - 30:
            pied = pied[:-2]
        if pied != propre_pdf(self.pied):
            pied = pied.rstrip() + "..."
        self.cell(0, 4, pied)
        self.set_x(self.l_margin)
        self.cell(0, 4, f"page {self.page_no()} / {{nb}}", align="R")


def resume_pieces(mail):
    lignes = []
    for p in mail.pieces:
        if p.genre == "technique":
            continue
        lignes.append(("[mail joint] " if p.dans else "") + decrire(p))
    for l in mail.liens:
        lab = f"{l['type']} (lien {l['hote']})"
        if l.get("piece") is not None:
            lignes.append(f"{lab} : téléchargé, " + decrire(l["piece"], court=True))
        elif l.get("statut"):
            lignes.append(f"{lab} : {l['statut']}")
    if mail.justificatifs:
        lignes.append(pluriel(mail.justificatifs, "justificatif") + " en lien non repris")
    return lignes or ["aucune"]


def decrire(p, court=False):
    nom = "" if court else p.nom + " : "
    if not p.imprimer:
        return nom + "non imprimée (" + ("jointe deux fois" if p.doublon_de else p.raison or "écartée") + ")"
    if p.genre == "image":
        return nom + "image, 1 page"
    deja = f" (même fichier que dans le mail du {date_courte(p.aussi_dans.date)})" if getattr(p, "aussi_dans", None) else ""
    if p.tronquer:
        return nom + f"{p.pages} pages, tronquée (pages 1 à {PAGES_DEBUT}, avis, page {p.pages})" + deja
    return nom + pluriel(p.pages, "page") + deja


def page_mail(mail, rang, total):
    pdf = Page(f"MAIL {rang} / {total}", date_courte(mail.date), mail.sujet)
    pdf.add_page()
    f = pdf.f
    pdf.set_text_color(*ENCRE)
    pdf.set_font(f, "B", 14)
    pdf.multi_cell(0, 6.3, propre_pdf(mail.sujet), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    def ligne(label, valeur, premier=True):
        pdf.set_font(f, "B", 9)
        pdf.set_text_color(*GRIS)
        pdf.cell(22, 4.6, propre_pdf(label) if premier else "")
        pdf.set_font(f, "", 9.5)
        pdf.set_text_color(*ENCRE)
        pdf.multi_cell(0, 4.6, propre_pdf(valeur), new_x="LMARGIN", new_y="NEXT")

    ligne("De", mail.de or "-")
    ligne("À", mail.a or "-")
    if mail.cc:
        ligne("Cc", mail.cc)
    ligne("Date", date_longue(mail.date))
    for i, r in enumerate(resume_pieces(mail)):
        ligne("Pièces", "• " + r, i == 0)
    pdf.ln(2.5)
    pdf.set_draw_color(*FILET)
    pdf.set_line_width(0.2)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(4)
    pdf.set_font(f, "", 9.5)
    for l in (mail.propre or "").split("\n"):
        if not l.strip():
            pdf.ln(2.4)
            continue
        pdf.set_text_color(*(GRIS if l.startswith(">") else ENCRE))
        pdf.multi_cell(0, 4.45, propre_pdf(l), new_x="LMARGIN", new_y="NEXT")
    return bytes(pdf.output())


def page_avis(mail, p):
    pdf = Page()
    pdf.add_page()
    f = pdf.f
    omises = p.pages - PAGES_DEBUT - 1
    x, w = 19, pdf.epw
    lignes = [
        ("B", 12, ENCRE, p.nom),
        ("", 10, GRIS, "Mail : " + mail.sujet),
        ("", 10, GRIS, "Envoyé le " + date_longue(mail.date)),
        None,
        ("", 11.5, ENCRE, f"Ce document compte {p.pages} pages."),
        ("", 11.5, ENCRE, f"Pages imprimées : 1 à {PAGES_DEBUT} et {p.pages} (dernière page)."),
        ("B", 11.5, ENCRE, f"Pages non imprimées : {PAGES_DEBUT + 1}" + (f" à {p.pages - 1}" if omises > 1 else "") + f", soit {pluriel(omises, 'page')}."),
        None,
        ("", 10.5, GRIS, f"Motif : {p.motif}."),
        ("", 10.5, GRIS, "Le document complet reste consultable depuis le mail d'origine."),
    ]
    y0 = 95
    pdf.set_y(y0 + 16)
    pdf.set_x(x + 10)
    pdf.set_font(f, "B", 24)
    pdf.set_text_color(*FRAMBOISE)
    pdf.cell(w - 20, 10, propre_pdf("DOCUMENT TRONQUÉ"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    for l in lignes:
        if l is None:
            pdf.ln(3.5)
            continue
        style, taille, coul, txt = l
        pdf.set_font(f, style, taille)
        pdf.set_text_color(*coul)
        pdf.set_x(x + 10)
        pdf.multi_cell(w - 20, taille * 0.5, propre_pdf(txt), new_x="LMARGIN", new_y="NEXT")
    y1 = pdf.get_y() + 8
    pdf.set_draw_color(*FRAMBOISE)
    pdf.set_line_width(0.5)
    pdf.rect(x, y0, w, y1 - y0)
    pdf.set_fill_color(*FRAMBOISE)
    pdf.rect(x, y0, w, 3, style="F")
    return bytes(pdf.output())


def page_image(p):
    from PIL import Image
    img = Image.open(io.BytesIO(p.data))
    img.load()
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    paysage = img.width > img.height * 1.15
    pdf = FPDF(orientation="L" if paysage else "P", format="A4", unit="mm")
    pdf.set_auto_page_break(False)
    pdf.add_page()
    mg = 10
    s = min((pdf.w - 2 * mg) / img.width, (pdf.h - 2 * mg) / img.height)
    w, h = img.width * s, img.height * s
    pdf.image(img, x=(pdf.w - w) / 2, y=mg, w=w, h=h)
    return bytes(pdf.output())


def ajouter_pdf(writer, data):
    r = PdfReader(io.BytesIO(data))
    for pg in r.pages:
        writer.add_page(pg)
    return len(r.pages)


# ----------------------------------------------------------------- traitement
class Resultat:
    def __init__(self):
        self.mails = []          # mails lisibles, classés par date
        self.a_imprimer = []     # mails retenus (hors doublons)
        self.doublons = []       # mails écartés car identiques à un mail précédent
        self.illisibles = []
        self.volumineux = []     # (mail, pièce) au-delà du seuil : troncature proposée
        self.a_verifier = []     # messages pour l'utilisateur


_compteur = [0]


def _uid():
    _compteur[0] += 1
    return f"p{_compteur[0]}"


def analyser(fichiers, progression=None):
    """fichiers : liste de (nom, octets) de mails .msg. Lit les mails, télécharge factures et RIB,
    repère doublons, factures Orange et documents volumineux. Rien n'est écrit sur disque."""
    res = Resultat()
    mails = [charger(nom, data) for nom, data in fichiers]
    res.illisibles = [m for m in mails if m.erreur]
    mails = [m for m in mails if not m.erreur]
    mails.sort(key=lambda m: (m.date is None, m.date.timestamp() if m.date else 0, m.chemin.name.lower()))
    vus_mail = {}
    for m in mails:
        if m.cle in vus_mail:
            m.doublon_de = vus_mail[m.cle]
        else:
            vus_mail[m.cle] = m
    res.mails = mails
    res.a_imprimer = [m for m in mails if not m.doublon_de]
    res.doublons = [m for m in mails if m.doublon_de]

    vus_piece = {}
    for i, m in enumerate(res.a_imprimer, 1):
        if progression:
            progression(i, len(res.a_imprimer), m)
        for p in m.pieces:
            p.analyser(m)
        if m.liens and not m.a_des_pieces():
            for l in m.liens:
                nom = l["fichier"] or f"{l['type']}.pdf"
                data, raison = telecharger(l["cible"])
                if data is None:
                    l["statut"] = "téléchargement impossible, " + raison
                    res.a_verifier.append(f"{m.sujet} : lien {l['type']} non récupéré ({raison}). Lien : {l['cible']}")
                    continue
                p = Piece(f"{l['type']} - {nom}", data, "lien")
                p.raison, p.imprimer = "", True
                p.analyser(m)
                l["piece"] = p
        for p in m.pieces + [l["piece"] for l in m.liens if l.get("piece")]:
            p.uid = _uid()
            if p.imprimer and p.genre in ("pdf", "image"):
                deja = vus_piece.get(p.empreinte)
                if deja is m:
                    p.doublon_de, p.imprimer = m, False
                elif deja is not None:
                    p.aussi_dans = deja
                    res.a_verifier.append(f"{m.sujet} : « {p.nom} » est identique à une pièce du mail du {date_courte(deja.date)} (imprimée deux fois)")
                else:
                    vus_piece[p.empreinte] = m
            if p.imprimer and p.genre == "pdf" and p.pages > PAGES_DEBUT + 1:
                if p.orange:
                    p.tronquer, p.motif = True, "facture Orange"
                elif p.pages > SEUIL_PAGES:
                    p.tronquer, p.motif = True, f"document volumineux (plus de {SEUIL_PAGES} pages)"
                    res.volumineux.append((m, p))
            if "à imprimer à part" in p.raison:
                res.a_verifier.append(f"{m.sujet} : « {p.nom} » {p.raison}")
        if not m.imprimees() and not m.liens:
            res.a_verifier.append(f"{m.sujet} : aucune pièce jointe ni lien de facture trouvé, seul le mail est imprimé")
    for m in res.illisibles:
        res.a_verifier.append(f"{m.chemin.name} : {m.erreur}")
    return res


def assembler(res, tronquer=None):
    """Construit le PDF. tronquer : {uid de pièce: bool} pour les documents volumineux.
    Renvoie (octets du PDF, nombre de pages, pièces non reprises)."""
    for _, p in res.volumineux:
        if tronquer is not None and p.uid in tronquer:
            p.tronquer = bool(tronquer[p.uid])
    w = PdfWriter()
    total, problemes = 0, []
    for i, m in enumerate(res.a_imprimer, 1):
        total += ajouter_pdf(w, page_mail(m, i, len(res.a_imprimer)))
        for p in m.imprimees():
            try:
                if p.genre == "image":
                    total += ajouter_pdf(w, page_image(p))
                    continue
                pages = p.lecteur.pages
                if p.tronquer:
                    for k in range(PAGES_DEBUT):
                        w.add_page(pages[k])
                    total += PAGES_DEBUT
                    total += ajouter_pdf(w, page_avis(m, p))
                    w.add_page(pages[p.pages - 1])
                    total += 1
                else:
                    for pg in pages:
                        w.add_page(pg)
                    total += p.pages
            except Exception:
                problemes.append(f"{m.sujet} : « {p.nom} »")
    try:
        w.add_metadata({"/Title": "Liasse mails et factures", "/Creator": "Liasse mails et factures (Capeos)"})
    except Exception:
        pass
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue(), total, problemes
