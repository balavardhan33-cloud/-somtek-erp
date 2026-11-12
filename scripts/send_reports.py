"""SEV7EN scheduled report mailer.

Runs hourly on GitHub Actions. Reads the report schedules saved in the ERP
(Reports > Scheduled Reports), builds each due report as an Excel file and
emails it to the schedule's recipients.

Secrets (GitHub repo > Settings > Secrets and variables > Actions):
  ERP_USERNAME   ERP admin login name (used to sign in to the database)
  ERP_PASSWORD   that login's password
  SMTP_USER      Gmail address the reports are sent from
  SMTP_PASS      Gmail app password (16 characters)
"""
import io, json, os, re, smtplib, sys, time, urllib.request, urllib.error
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

from openpyxl import Workbook

API_KEY = "AIzaSyAapUWQ9Mm1Q7bBUDG1avriThBA_YqMYfE"
PROJECT = "somtek-erp"
COLLECTION = "somtek_erp"
AUTH_DOMAIN = "somtek-erp.local"
BASE = f"https://firestore.googleapis.com/v1/projects/{PROJECT}/databases/(default)/documents/{COLLECTION}/"
PARTS_TAG = "SEV7EN_PARTS|"
IST = timezone(timedelta(hours=5, minutes=30))
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

TOKEN = ""


def http(url, data=None, method=None, headers=None):
    h = dict(headers or {})
    if TOKEN:
        h["Authorization"] = "Bearer " + TOKEN
    body = None
    if data is not None:
        body = json.dumps(data).encode()
        h["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=body, method=method, headers=h)
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode() or "{}")


def sign_in():
    global TOKEN
    user = os.environ.get("ERP_USERNAME", "").strip().lower().replace(" ", "")
    pw = os.environ.get("ERP_PASSWORD", "")
    if not user or not pw:
        print("No ERP_USERNAME/ERP_PASSWORD set - reading without sign-in")
        return
    url = "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=" + API_KEY
    try:
        j = http(url, {"email": f"{user}@{AUTH_DOMAIN}", "password": pw, "returnSecureToken": True})
        TOKEN = j.get("idToken", "")
        print("Signed in to the database as", user)
    except urllib.error.HTTPError as e:
        print("Database sign-in failed:", e.read().decode()[:300], "- trying without sign-in")


def read_raw(doc_id):
    try:
        j = http(BASE + doc_id)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    return ((j.get("fields") or {}).get("data") or {}).get("stringValue")


def fs_get(doc_id):
    """Same format as the ERP's fsGet: plain JSON, or a pointer to parts."""
    for attempt in range(3):
        raw = read_raw(doc_id)
        if raw is None:
            return None
        if not raw.startswith(PARTS_TAG):
            return json.loads(raw)
        _, slot, n, stamp = raw.split("|")[:4]
        head = stamp + "|"
        out, ok = [], True
        for i in range(int(n)):
            part = read_raw(f"{doc_id}__{slot}{i}")
            if not part or not part.startswith(head):
                ok = False
                break
            out.append(part[len(head):])
        if ok:
            return json.loads("".join(out))
        time.sleep(1 + attempt)
    raise RuntimeError("parts of " + doc_id + " changed while reading")


def fs_set_small(doc_id, data):
    http(BASE + doc_id, {"fields": {"data": {"stringValue": json.dumps(data)}}}, method="PATCH")


# ---------- helpers ported from the ERP ----------
def dparse(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if "T" in s:
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(IST).replace(tzinfo=None)
        except ValueError:
            return None
    m = re.match(r"^(\d{1,4})[.\-/](\d{1,2})[.\-/](\d{1,4})$", s)
    if not m:
        return None
    a, b, c = m.groups()
    y, mo, d = (int(a), int(b), int(c)) if len(a) == 4 else (int(c), int(b), int(a))
    if y < 100:
        y += 2000
    try:
        return datetime(y, mo, d)
    except ValueError:
        return None


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def money(v):
    return round(num(v) * 100) / 100


def today_str():
    return datetime.now(IST).strftime("%Y-%m-%d")


def sales_balance(r):
    try:
        inv = float(r.get("invoiceValue"))
    except (TypeError, ValueError):
        return num(r.get("balanceOutstanding"))
    return round(max(0.0, inv - num(r.get("amountReceived"))) * 100) / 100


def overdue(r):
    try:
        days = int(r.get("creditDays"))
    except (TypeError, ValueError):
        return False
    d = dparse(r.get("date"))
    if not d:
        return False
    due = (d + timedelta(days=days)).strftime("%Y-%m-%d")
    return (r.get("paymentStatus") or "Pending") == "Pending" and today_str() > due


def inv_age_days(i, d):
    nm = str(i.get("name") or "").strip().lower()
    ts = None
    if nm:
        for r in d.get("inward") or []:
            if str(r.get("item") or "").strip().lower() == nm:
                t = dparse(r.get("date"))
                if t and (ts is None or t > ts):
                    ts = t
    if ts is None:
        ts = dparse(i.get("lastInwardOn") or i.get("addedOn") or "")
    if ts is None:
        return None
    return max(0, (datetime.now(IST).replace(tzinfo=None) - ts).days)


def low_stock(i):
    return i.get("deadStock") == "Yes" or num(i.get("qty")) == 0


def dead_stock(i, d):
    if num(i.get("qty")) <= 0:
        return False
    a = inv_age_days(i, d)
    return a is not None and a > 730


def age_label(days):
    if days is None or days <= 90:
        return ""
    if days <= 365:
        return ">90 days"
    if days <= 730:
        return ">1 yr"
    return ">" + str(days // 365) + " yrs"


def in_period(v, fr, to):
    d = date_only(v)
    return bool(d) and (not fr or d >= fr) and (not to or d <= to)


def sheet_for(section, loc, d, schema, period=""):
    if section == "summary":
        sa, pa, inv = d.get("salesAccounts") or [], d.get("purchaseAccounts") or [], d.get("inventory") or []
        pending = lambda r: (r.get("paymentStatus") or "Pending") == "Pending"
        rows = [
            ["Report generated", datetime.now(IST).strftime("%d-%m-%Y %I:%M %p")],
            ["Location", loc],
            ["Total Sales (Invoice Value)", money(sum(num(r.get("invoiceValue")) for r in sa))],
            ["Sales Outstanding", money(sum(sales_balance(r) for r in sa))],
            ["Sales Payments Pending", sum(1 for r in sa if pending(r))],
            ["Sales Payments Overdue", sum(1 for r in sa if overdue(r))],
            ["Total Purchase Value", money(sum(num(r.get("grossTotal")) for r in pa))],
            ["Purchase Outstanding", money(sum(num(r.get("balanceOutstanding")) for r in pa))],
            ["Purchase Payments Overdue", sum(1 for r in pa if overdue(r))],
            ["Inventory Items", len(inv)],
            ["Inventory Value", money(sum(num(i.get("qty")) * num(i.get("cost")) for i in inv))],
            ["Aging Stock (>90 days, <2 yrs)", sum(1 for i in inv if not low_stock(i) and not dead_stock(i, d) and (inv_age_days(i, d) or 0) > 90)],
            ["Low Stock Items", sum(1 for i in inv if low_stock(i))],
            ["Dead Stock Items (2 yrs+)", sum(1 for i in inv if dead_stock(i, d))],
            ["Pending Enquiries", sum(1 for r in d.get("enquiry") or [] if (r.get("action") or "Pending") == "Pending")],
        ]
        return loc + " Summary", ["Metric", "Value"], rows
    if section == "outstanding":
        rows = []
        for r in d.get("salesAccounts") or []:
            if (r.get("paymentStatus") or "Pending") == "Pending":
                rows.append(["SALE", r.get("date") or "", r.get("invNo") or "", r.get("customer") or "",
                             money(r.get("invoiceValue")), sales_balance(r), "OVERDUE" if overdue(r) else "Pending"])
        for r in d.get("purchaseAccounts") or []:
            if (r.get("paymentStatus") or "Pending") == "Pending":
                rows.append(["PURCHASE", r.get("date") or "", r.get("invoiceNo") or r.get("billNo") or "",
                             r.get("vendor") or r.get("supplier") or "", money(r.get("grossTotal")),
                             money(r.get("balanceOutstanding")), "OVERDUE" if overdue(r) else "Pending"])
        return loc + " Outstanding", ["Type", "Date", "Invoice No", "Party", "Amount", "Balance", "Status"], rows
    if section == "inventory":
        rows = []
        for i in d.get("inventory") or []:
            days = inv_age_days(i, d)
            flag = "LOW STOCK" if low_stock(i) else ("DEAD STOCK" if dead_stock(i, d) else age_label(days))
            rows.append([i.get("name"), i.get("qty"), i.get("cost"), (num(i.get("qty")) * num(i.get("cost"))) or "",
                         i.get("category") or "", i.get("lastInwardOn") or i.get("addedOn") or "",
                         "" if days is None else days, flag])
        return loc + " Inventory", ["Item", "Qty", "Cost", "Total Value", "Category", "Stock Since", "Age (days)", "Flag"], rows
    sec = (schema.get("sections") or {}).get(section) or {}
    keys, headers = sec.get("keys") or [], sec.get("headers") or []
    label = sec.get("label") or section
    fr, to = period_range(period, {}) if period and period != "all" and sec.get("dateKey") else ("", "")
    rows = []
    for r in d.get(section) or []:
        if isinstance(r, dict) and (not (fr or to) or in_period(r.get(sec["dateKey"]), fr, to)):
            rows.append([cell(r.get(k)) for k in keys] + [r.get("addedBy") or ""])
    return loc + " " + label, headers + ["Added By"], rows


def date_only(v):
    """YYYY-MM-DD for a stored date (also YYYY-MM month values), else ''."""
    s = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})$", s)
    if m:
        return s + "-01"
    d = dparse(s)
    return d.strftime("%Y-%m-%d") if d else ""


def period_range(p, b):
    now = datetime.now(IST).replace(tzinfo=None)
    iso = lambda d: d.strftime("%Y-%m-%d")
    if p in ("lw", "tw"):
        mon = now - timedelta(days=now.weekday() + (7 if p == "lw" else 0))
        return iso(mon), iso(mon + timedelta(days=5))
    if p == "yday":
        d = now - timedelta(days=1)
        return iso(d), iso(d)
    if p == "d7":
        return iso(now - timedelta(days=6)), iso(now)
    if p == "d30":
        return iso(now - timedelta(days=29)), iso(now)
    if p == "tm":
        first = now.replace(day=1)
        nxt = (first + timedelta(days=32)).replace(day=1)
        return iso(first), iso(nxt - timedelta(days=1))
    if p == "lm":
        first_this = now.replace(day=1)
        last = first_this - timedelta(days=1)
        return iso(last.replace(day=1)), iso(last)
    if p == "fy":
        y = now.year if now.month >= 4 else now.year - 1
        return f"{y}-04-01", f"{y + 1}-03-31"
    if p == "fixed":
        return b.get("from") or "", b.get("to") or ""
    return "", ""


def builder_sheet(b, loc, d):
    """Same rules as the ERP Report Builder."""
    cols = b.get("cols") or []
    rows = [r for r in (d.get(b.get("src")) or []) if isinstance(r, dict)]
    df = b.get("dateField") or ""
    if df:
        fr, to = period_range(b.get("period") or "all", b)
        if fr or to:
            def keep(r):
                v = date_only(r.get(df))
                return bool(v) and (not fr or v >= fr) and (not to or v <= to)
            rows = [r for r in rows if keep(r)]
    fk, fv = b.get("fKey") or "", str(b.get("fVal") or "").strip().lower()
    if fk and fv:
        rows = [r for r in rows if fv in str(r.get(fk) if r.get(fk) is not None else "").lower()]
    q = str(b.get("search") or "").strip().lower()
    if q:
        rows = [r for r in rows if any(q in str(r.get(c["key"]) if r.get(c["key"]) is not None else "").lower() for c in cols)]
    sk = b.get("sortKey") or df
    if sk:
        is_num = any(c["key"] == sk and c.get("type") == "number" for c in cols)
        if is_num:
            key = lambda r: num(r.get(sk))
        elif sk == df:
            key = lambda r: date_only(r.get(sk))
        else:
            key = lambda r: str(r.get(sk) if r.get(sk) is not None else "")
        rows.sort(key=key, reverse=(b.get("sortDir") != "asc"))
    return (loc + " " + (b.get("label") or b.get("src") or "Report"),
            [c["label"] for c in cols],
            [[cell(r.get(c["key"])) for c in cols] for r in rows])


def cell(v):
    if v is None:
        return ""
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def build_workbook(s, schema, loc_cache):
    locs = (schema.get("locations") or ["SOMTEK", "EVISSAT"]) if s.get("loc") == "Both" else [s.get("loc")]
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for loc in locs:
        if loc not in loc_cache:
            loc_cache[loc] = fs_get("loc_" + loc) or {}
        d = loc_cache[loc]
        jobs = [sheet_for(sec, loc, d, schema, s.get("period") or "") for sec in s.get("sections") or []]
        if s.get("builder"):
            jobs.append(builder_sheet(s["builder"], loc, d))
        for name, headers, rows in jobs:
            nm = re.sub(r"[\\/?*\[\]:]", " ", name)[:31] or "Sheet"
            base, k = nm, 2
            while nm in used:
                nm = base[:28] + " " + str(k)
                k += 1
            used.add(nm)
            ws = wb.create_sheet(nm)
            ws.append(headers)
            for r in rows:
                ws.append([c if isinstance(c, (int, float)) else str(c) for c in r])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def current_slot(s, now):
    """Send time for the current period (IST, naive)."""
    hour = int(s.get("sendHour") or 9)
    base = now.replace(hour=hour, minute=0, second=0, microsecond=0)
    freq = s.get("freq") or "Weekly"
    if freq == "Daily":
        return base
    if freq == "Weekly":
        monday = base - timedelta(days=base.weekday())
        return monday + timedelta(days=int(s.get("sendWeekday") or 0))
    return base.replace(day=min(28, max(1, int(s.get("sendMonthday") or 1))))


def is_due(s, now):
    if not any((r or {}).get("email") for r in s.get("recipients") or []):
        return False
    if s.get("sendNow"):
        return True
    slot = current_slot(s, now)
    if now < slot:
        return False
    last = dparse(s.get("lastSent")) if s.get("lastSent") else None
    return last is None or last < slot


def send_mail(to_list, subject, body, filename, data):
    user = os.environ.get("SMTP_USER", "").strip()
    pw = os.environ.get("SMTP_PASS", "").replace(" ", "").strip()
    if not user or not pw:
        raise RuntimeError("SMTP_USER / SMTP_PASS secret is empty - add them in GitHub Settings > Secrets")
    if len(pw) != 16:
        print(f"Note: SMTP_PASS has {len(pw)} characters; a Gmail app password has 16")
    msg = EmailMessage()
    msg["From"] = "SEV7EN Reports <" + user + ">"
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(data, maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)
    errors = []
    # try port 587 (STARTTLS) first, then 465 (SSL); retry once each
    for attempt in range(2):
        for port in (587, 465):
            stage = "connect"
            try:
                if port == 587:
                    smtp = smtplib.SMTP("smtp.gmail.com", 587, timeout=60)
                    smtp.ehlo()
                    stage = "starttls"
                    smtp.starttls()
                    smtp.ehlo()
                else:
                    smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60)
                with smtp:
                    stage = "login"
                    smtp.login(user, pw)
                    stage = "send"
                    smtp.send_message(msg)
                return
            except smtplib.SMTPAuthenticationError as e:
                raise RuntimeError(f"Gmail rejected the login ({e.smtp_code}) - check SMTP_USER and the 16-letter app password in SMTP_PASS")
            except Exception as e:
                errors.append(f"port {port} {stage}: {e}")
                print("  mail attempt failed -", errors[-1])
        time.sleep(5)
    raise RuntimeError("; ".join(errors[-2:]))


def main():
    sign_in()
    sched = fs_get("report_schedules") or {"list": []}
    schema = fs_get("report_schema") or {}
    now = datetime.now(IST).replace(tzinfo=None)
    due = [s for s in sched.get("list") or [] if is_due(s, now)]
    print(f"{len(sched.get('list') or [])} schedule(s), {len(due)} due at {now:%d-%m-%Y %H:%M} IST")
    if not due:
        return 0
    results, cache = {}, {}
    for s in due:
        to = sorted({r["email"].strip() for r in s.get("recipients") or [] if r and r.get("email")})
        stamp = now.strftime("%d-%m-%Y")
        fname = re.sub(r"[^A-Za-z0-9_\-]+", "_", s.get("name") or "SEV7EN_Report") + "_" + now.strftime("%Y-%m-%d") + ".xlsx"
        try:
            data = build_workbook(s, schema, cache)
            secs = ", ".join((schema.get("sections") or {}).get(k, {}).get("label", k) for k in s.get("sections") or [])
            if s.get("builder"):
                secs = ", ".join(filter(None, [secs, (s["builder"].get("label") or "") + " (Report Builder)"]))
            per = (s.get("builder") or {}).get("period") if s.get("builder") else s.get("period")
            per_txt = ""
            if per and per != "all":
                fr, to = period_range(per, s.get("builder") or {})
                if fr:
                    fmt = lambda x: datetime.strptime(x, "%Y-%m-%d").strftime("%d-%m-%Y")
                    per_txt = f"\nPeriod: {fmt(fr)} to {fmt(to)}"
            body = (f"Hello,\n\nPlease find attached the {s.get('freq', '').lower()} report \"{s.get('name')}\" "
                    f"({s.get('loc')}) generated on {stamp}.{per_txt}\n\nSections: {secs}\n\n- SEV7EN ERP (automatic email)\n")
            send_mail(to, f"[SEV7EN] {s.get('name')} - {stamp}", body, fname, data)
            results[s["id"]] = "ok"
            print("Sent", s.get("name"), "to", ", ".join(to))
        except Exception as e:  # keep going with the other schedules
            results[s["id"]] = "Email failed: " + str(e)[:120]
            print("FAILED", s.get("name"), e)
    # write the results back on a fresh copy, so edits made meanwhile are kept
    latest = fs_get("report_schedules") or {"list": []}
    sent_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    for s in latest.get("list") or []:
        if s.get("id") in results:
            s["lastStatus"] = results[s["id"]]
            s["sendNow"] = False
            if results[s["id"]] == "ok":
                s["lastSent"] = sent_at
    fs_set_small("report_schedules", latest)
    return 1 if any(v != "ok" for v in results.values()) else 0


if __name__ == "__main__":
    sys.exit(main())
