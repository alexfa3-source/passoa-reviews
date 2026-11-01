#!/usr/bin/env python3
"""
Email Unsubscriber
------------------
Σαρώνει το γραμματοκιβώτιό σου (IMAP), βρίσκει τα διαφημιστικά / newsletter
e-mail που σου έρχονται ακόμα (ενεργοί αποστολείς) και σου δίνει την επιλογή
να κάνεις unsubscribe από όσους θέλεις.

Χρησιμοποιεί μόνο τη standard library της Python (3.8+).

Πώς αναγνωρίζεται ένα διαφημιστικό e-mail:
  - Έχει header "List-Unsubscribe" (όλα τα σοβαρά newsletter το έχουν, και
    το Gmail/Yahoo το απαιτούν από τους μαζικούς αποστολείς).

Μέθοδοι unsubscribe (με σειρά προτίμησης):
  1. One-click (RFC 8058): αυτόματο HTTPS POST — γίνεται αμέσως, χωρίς browser.
  2. mailto: στέλνει αυτόματα e-mail διαγραφής μέσω SMTP.
  3. Σύνδεσμος (http/https): ανοίγει στον browser για να το ολοκληρώσεις.
"""

import argparse
import datetime as dt
import email
import getpass
import imaplib
import json
import os
import re
import smtplib
import ssl
import sys
import urllib.parse
import urllib.request
import webbrowser
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr, parsedate_to_datetime

HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unsubscribed.json")

HEADER_FIELDS = (
    "FROM SUBJECT DATE LIST-UNSUBSCRIBE LIST-UNSUBSCRIBE-POST LIST-ID PRECEDENCE"
)

# Γνωστοί πάροχοι: (IMAP host, SMTP host)
PROVIDERS = {
    "gmail.com": ("imap.gmail.com", "smtp.gmail.com"),
    "googlemail.com": ("imap.gmail.com", "smtp.gmail.com"),
    "outlook.com": ("outlook.office365.com", "smtp.office365.com"),
    "hotmail.com": ("outlook.office365.com", "smtp.office365.com"),
    "live.com": ("outlook.office365.com", "smtp.office365.com"),
    "yahoo.com": ("imap.mail.yahoo.com", "smtp.mail.yahoo.com"),
    "icloud.com": ("imap.mail.me.com", "smtp.mail.me.com"),
}


# --------------------------------------------------------------------------- #
# Βοηθητικές συναρτήσεις ανάλυσης headers
# --------------------------------------------------------------------------- #
def decode_str(value):
    """Αποκωδικοποιεί encoded headers (π.χ. =?UTF-8?B?...?=) σε κανονικό κείμενο."""
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return str(value)


def clean_ws(text):
    return re.sub(r"\s+", " ", text or "").strip()


def parse_list_unsubscribe(value):
    """Επιστρέφει (λίστα https/http URLs, λίστα mailto URLs)."""
    links = re.findall(r"<\s*([^>]+?)\s*>", clean_ws(value))
    if not links and value:
        # Κάποιοι αποστολείς ξεχνούν τα < >
        links = [p.strip() for p in clean_ws(value).split(",") if p.strip()]
    http = [l for l in links if l.lower().startswith(("http://", "https://"))]
    mailto = [l for l in links if l.lower().startswith("mailto:")]
    return http, mailto


def parse_date(value):
    try:
        d = parsedate_to_datetime(value)
        if d.tzinfo is None:
            d = d.replace(tzinfo=dt.timezone.utc)
        return d
    except Exception:
        return None


class Sender:
    """Ένας αποστολέας διαφημιστικών με όλα τα στοιχεία που μαζέψαμε."""

    def __init__(self, address, name):
        self.address = address
        self.name = name
        self.count = 0
        self.last_date = None
        self.last_subject = ""
        self.http_links = []
        self.mailto_links = []
        self.one_click = False

    def add(self, date, subject, http, mailto, one_click):
        self.count += 1
        # Κρατάμε τα στοιχεία του πιο πρόσφατου e-mail (τα links μπορεί να λήγουν)
        if self.last_date is None or (date and date > self.last_date):
            self.last_date = date
            self.last_subject = subject
            if http or mailto:
                self.http_links = http
                self.mailto_links = mailto
                self.one_click = one_click
        elif not (self.http_links or self.mailto_links):
            self.http_links, self.mailto_links, self.one_click = http, mailto, one_click

    @property
    def method(self):
        if self.one_click and any(l.lower().startswith("https://") for l in self.http_links):
            return "one-click"
        if self.mailto_links:
            return "e-mail"
        if self.http_links:
            return "browser"
        return "-"


def analyze_headers(raw_headers_list):
    """Παίρνει λίστα από raw headers (bytes) και επιστρέφει dict address -> Sender."""
    senders = {}
    for raw in raw_headers_list:
        msg = email.message_from_bytes(raw)
        lu = msg.get("List-Unsubscribe")
        if not lu:
            continue
        name, addr = parseaddr(decode_str(msg.get("From", "")))
        addr = addr.lower().strip()
        if not addr:
            continue
        http, mailto = parse_list_unsubscribe(decode_str(lu))
        if not http and not mailto:
            continue
        one_click = "one-click" in clean_ws(msg.get("List-Unsubscribe-Post", "")).lower()
        date = parse_date(msg.get("Date", ""))
        subject = clean_ws(decode_str(msg.get("Subject", "")))
        s = senders.get(addr)
        if s is None:
            s = senders[addr] = Sender(addr, clean_ws(name) or addr)
        s.add(date, subject, http, mailto, one_click)
    return senders


# --------------------------------------------------------------------------- #
# IMAP
# --------------------------------------------------------------------------- #
def fetch_headers(imap_host, user, password, folder, days, limit):
    print(f"\n🔌 Σύνδεση στο {imap_host} ως {user} ...")
    imap = imaplib.IMAP4_SSL(imap_host)
    imap.login(user, password)
    typ, _ = imap.select(f'"{folder}"', readonly=True)  # readonly: δεν αλλάζει τίποτα (π.χ. "αναγνωσμένα")
    if typ != "OK":
        raise SystemExit(f"❌ Δεν βρέθηκε ο φάκελος '{folder}'.")

    since = (dt.date.today() - dt.timedelta(days=days)).strftime("%d-%b-%Y")
    typ, data = imap.uid("search", None, "SINCE", since)
    uids = data[0].split() if data and data[0] else []
    if limit and len(uids) > limit:
        uids = uids[-limit:]  # τα πιο πρόσφατα
    print(f"📬 Βρέθηκαν {len(uids)} e-mail τις τελευταίες {days} ημέρες. Ανάλυση headers...")

    raws = []
    chunk = 300
    for i in range(0, len(uids), chunk):
        batch = b",".join(uids[i:i + chunk]).decode()
        typ, resp = imap.uid("fetch", batch, f"(BODY.PEEK[HEADER.FIELDS ({HEADER_FIELDS})])")
        if typ != "OK":
            continue
        for part in resp:
            if isinstance(part, tuple) and len(part) == 2:
                raws.append(part[1])
        done = min(i + chunk, len(uids))
        print(f"   ... {done}/{len(uids)}", end="\r")
    print()
    imap.logout()
    return raws


# --------------------------------------------------------------------------- #
# Unsubscribe
# --------------------------------------------------------------------------- #
def unsubscribe_one_click(url):
    req = urllib.request.Request(
        url,
        data=b"List-Unsubscribe=One-Click",
        method="POST",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "User-Agent": "Mozilla/5.0 (EmailUnsubscriber)",
        },
    )
    with urllib.request.urlopen(req, timeout=20, context=ssl.create_default_context()) as r:
        return 200 <= r.status < 300


def unsubscribe_mailto(mailto, smtp_host, user, password):
    parsed = urllib.parse.urlparse(mailto)
    to = urllib.parse.unquote(parsed.path)
    q = urllib.parse.parse_qs(parsed.query)
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to
    msg["Subject"] = q.get("subject", ["unsubscribe"])[0]
    msg.set_content(q.get("body", ["unsubscribe"])[0])
    ctx = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL(smtp_host, 465, context=ctx, timeout=30) as s:
            s.login(user, password)
            s.send_message(msg)
    except (smtplib.SMTPException, OSError):
        with smtplib.SMTP(smtp_host, 587, timeout=30) as s:
            s.starttls(context=ctx)
            s.login(user, password)
            s.send_message(msg)
    return True


def do_unsubscribe(sender, smtp_host, user, password, dry_run):
    label = f"{sender.name} <{sender.address}>"
    if dry_run:
        print(f"   [dry-run] {label}: θα γινόταν μέσω '{sender.method}'")
        return "dry-run"

    # 1) One-click
    if sender.method == "one-click":
        url = next(l for l in sender.http_links if l.lower().startswith("https://"))
        try:
            if unsubscribe_one_click(url):
                print(f"   ✅ {label}: έγινε unsubscribe (one-click)")
                return "one-click"
        except Exception as e:
            print(f"   ⚠️  {label}: το one-click απέτυχε ({e}), δοκιμάζω άλλη μέθοδο...")

    # 2) mailto
    if sender.mailto_links:
        try:
            unsubscribe_mailto(sender.mailto_links[0], smtp_host, user, password)
            print(f"   ✅ {label}: στάλθηκε e-mail διαγραφής")
            return "e-mail"
        except Exception as e:
            print(f"   ⚠️  {label}: η αποστολή e-mail απέτυχε ({e})")

    # 3) Browser
    if sender.http_links:
        url = sender.http_links[0]
        print(f"   🌐 {label}: ανοίγει στον browser — ολοκλήρωσε εκεί:\n      {url}")
        try:
            webbrowser.open(url)
        except Exception:
            pass
        return "browser"

    print(f"   ❌ {label}: δεν βρέθηκε τρόπος unsubscribe")
    return None


# --------------------------------------------------------------------------- #
# Ιστορικό
# --------------------------------------------------------------------------- #
def load_history():
    try:
        with open(HISTORY_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_history(history):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
def parse_selection(text, n):
    """'1,3,5-7' -> {1,3,5,6,7}. 'all' -> όλα."""
    text = text.strip().lower()
    if text in ("all", "όλα", "ola", "*"):
        return set(range(1, n + 1))
    chosen = set()
    for part in re.split(r"[,\s]+", text):
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            if a.isdigit() and b.isdigit():
                chosen.update(range(int(a), int(b) + 1))
        elif part.isdigit():
            chosen.add(int(part))
    return {i for i in chosen if 1 <= i <= n}


def truncate(text, width):
    return text if len(text) <= width else text[: width - 1] + "…"


def print_table(rows, history):
    print()
    print(f"{'#':>3}  {'Αποστολέας':<34} {'Πλήθος':>6}  {'Τελευταίο':<10}  {'Μέθοδος':<9}  Τελευταίο θέμα")
    print("-" * 110)
    for i, s in enumerate(rows, 1):
        last = s.last_date.strftime("%d/%m/%Y") if s.last_date else "?"
        flag = ""
        if s.address in history:
            flag = "  ⚠️ ξαναστέλνει μετά το unsubscribe!"
            if s.last_date and history[s.address].get("date"):
                try:
                    if s.last_date <= dt.datetime.fromisoformat(history[s.address]["date"]):
                        flag = "  (έχει γίνει unsubscribe)"
                except ValueError:
                    pass
        who = truncate(f"{s.name} <{s.address}>" if s.name != s.address else s.address, 34)
        print(f"{i:>3}  {who:<34} {s.count:>6}  {last:<10}  {s.method:<9}  {truncate(s.last_subject, 40)}{flag}")
    print()


def main():
    ap = argparse.ArgumentParser(description="Βρίσκει διαφημιστικά e-mail και κάνει unsubscribe.")
    ap.add_argument("--email", default=os.environ.get("EMAIL_USER"), help="Η διεύθυνσή σου (ή env EMAIL_USER)")
    ap.add_argument("--imap", help="IMAP server (αυτόματο για Gmail/Outlook/Yahoo/iCloud)")
    ap.add_argument("--smtp", help="SMTP server (για unsubscribe μέσω e-mail)")
    ap.add_argument("--folder", default="INBOX", help='Φάκελος (π.χ. INBOX ή "[Gmail]/All Mail")')
    ap.add_argument("--days", type=int, default=60, help="Πόσες μέρες πίσω να ψάξει (default 60)")
    ap.add_argument("--min-count", type=int, default=1, help="Δείξε μόνο αποστολείς με ≥ N e-mail")
    ap.add_argument("--limit", type=int, default=5000, help="Μέγιστο πλήθος e-mail προς ανάλυση")
    ap.add_argument("--dry-run", action="store_true", help="Δείξε τι θα γινόταν χωρίς να κάνει unsubscribe")
    args = ap.parse_args()

    user = args.email or input("📧 E-mail: ").strip()
    domain = user.split("@")[-1].lower()
    imap_host, smtp_host = PROVIDERS.get(domain, (None, None))
    imap_host = args.imap or imap_host or f"imap.{domain}"
    smtp_host = args.smtp or smtp_host or f"smtp.{domain}"

    password = os.environ.get("EMAIL_APP_PASSWORD") or getpass.getpass(
        "🔑 App password (δεν εμφανίζεται όταν το γράφεις): "
    )

    try:
        raws = fetch_headers(imap_host, user, password, args.folder, args.days, args.limit)
    except imaplib.IMAP4.error as e:
        sys.exit(f"❌ Αποτυχία σύνδεσης: {e}\n   Για Gmail χρειάζεσαι App Password (δες README.md).")

    senders = analyze_headers(raws)
    rows = sorted(
        (s for s in senders.values() if s.count >= args.min_count),
        key=lambda s: (s.count, s.last_date or dt.datetime.min.replace(tzinfo=dt.timezone.utc)),
        reverse=True,
    )
    if not rows:
        print("🎉 Δεν βρέθηκαν ενεργά διαφημιστικά e-mail!")
        return

    history = load_history()
    print(f"\n📢 Βρέθηκαν {len(rows)} ενεργοί αποστολείς διαφημιστικών/newsletter "
          f"(από {sum(s.count for s in rows)} e-mail):")
    print_table(rows, history)

    print("Διάλεξε από ποιους θέλεις unsubscribe:")
    print("  π.χ. 1,3,5-8   ή   all   ή   Enter για έξοδο")
    choice = input("> ").strip()
    selected = parse_selection(choice, len(rows)) if choice else set()
    if not selected:
        print("👋 Δεν επιλέχθηκε τίποτα. Έξοδος.")
        return

    chosen = [rows[i - 1] for i in sorted(selected)]
    print("\nΘα γίνει unsubscribe από:")
    for s in chosen:
        print(f"   • {s.name} <{s.address}>  [{s.method}]")
    if input("\nΣυνέχεια; (ν/ο) > ").strip().lower() not in ("ν", "ναι", "n", "nai", "y", "yes"):
        print("Ακυρώθηκε.")
        return

    print()
    stats = {}
    for s in chosen:
        result = do_unsubscribe(s, smtp_host, user, password, args.dry_run)
        stats[result] = stats.get(result, 0) + 1
        if result and result != "dry-run":
            history[s.address] = {
                "name": s.name,
                "method": result,
                "date": dt.datetime.now(dt.timezone.utc).isoformat(),
            }
    if not args.dry_run:
        save_history(history)

    print("\n📊 Σύνοψη:")
    names = {"one-click": "Αυτόματα (one-click)", "e-mail": "Μέσω e-mail",
             "browser": "Άνοιξαν στον browser", "dry-run": "Dry-run", None: "Απέτυχαν"}
    for k, v in stats.items():
        print(f"   {names.get(k, k)}: {v}")
    print("\nΣημ.: Κάποιοι αποστολείς χρειάζονται έως ~10 μέρες για να σταματήσουν. "
          "Αν ξανατρέξεις το πρόγραμμα, όσοι συνεχίζουν θα σημειώνονται με ⚠️.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nΔιακόπηκε.")
