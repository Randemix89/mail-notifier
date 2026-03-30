import csv
import mimetypes
import os
import ssl
import threading
import time
import traceback
import json
import random
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    TOP,
    VERTICAL,
    W,
    Button,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    Listbox,
    Menu,
    Message,
    Scrollbar,
    StringVar,
    Text,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
)
from tkinter import ttk
from typing import Iterable, List, Optional

import smtplib
import re
from email.utils import parseaddr, formataddr
import webbrowser
import socket
import time as _time

try:
    import socks  # type: ignore
except Exception:
    socks = None

try:
    import certifi  # type: ignore
except Exception:
    certifi = None

try:
    import sv_ttk  # type: ignore
except Exception:
    sv_ttk = None

APP_NAME = "Mail Notifier"
APP_REPO_URL = "https://github.com/Randemix89/mail-notifier"

def app_version() -> str:
    try:
        p = os.path.join(APP_DIR, "VERSION")
        with open(p, "r", encoding="utf-8") as f:
            v = (f.read() or "").strip()
        return v or "0.0.0"
    except Exception:
        return "0.0.0"


PROVIDERS = {
    "Gmail": {"host": "smtp.gmail.com", "port": 587, "starttls": 1},
    "Mail.ru": {"host": "smtp.mail.ru", "port": 587, "starttls": 1},
    "Yandex": {"host": "smtp.yandex.ru", "port": 587, "starttls": 1},
    "Rambler": {"host": "smtp.rambler.ru", "port": 587, "starttls": 1},
    "GMX": {"host": "smtp.gmx.com", "port": 587, "starttls": 1},
    "Custom": {"host": "", "port": 587, "starttls": 1},
}

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")


def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _guess_delimiter(first_line: str) -> Optional[str]:
    common_delims = [",", ";", "\t", "|"]
    return next((d for d in common_delims if d in (first_line or "")), None)


def _provider_defaults_for_username(username: str) -> dict:
    u = (username or "").strip().lower()
    # Common providers (used for auto-fill on import)
    if u.endswith("@gmail.com"):
        return {"provider": "Gmail", **PROVIDERS["Gmail"]}
    if u.endswith("@mail.ru") or u.endswith("@inbox.ru") or u.endswith("@bk.ru") or u.endswith("@list.ru"):
        return {"provider": "Mail.ru", **PROVIDERS["Mail.ru"]}
    if u.endswith("@yandex.ru") or u.endswith("@ya.ru") or u.endswith("@yandex.com"):
        return {"provider": "Yandex", **PROVIDERS["Yandex"]}
    if u.endswith("@rambler.ru") or u.endswith("@ro.ru") or u.endswith("@lenta.ru") or u.endswith("@autorambler.ru"):
        return {"provider": "Rambler", **PROVIDERS["Rambler"]}
    if u.endswith("@gmx.com") or u.endswith("@gmx.de") or u.endswith("@gmx.net") or u.endswith("@gmx.at") or u.endswith("@gmx.ch"):
        return {"provider": "GMX", **PROVIDERS["GMX"]}
    return {"provider": "Custom", **PROVIDERS["Custom"]}


def read_accounts_from_file(path: str) -> List[dict]:
    """
    Import SMTP accounts from a file.

    Supported formats:
    - Lines: "email:password" or "email;password" or "email,password" or "email password"
    - CSV/TSV with headers (case-insensitive):
        username/email, password/pass, name, provider, host, port, starttls, verify_tls, from_email, sender_name
    """
    out: List[dict] = []
    if not path:
        return out

    # JSON: allow exporting/importing config snippets
    if path.lower().endswith(".json"):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            items = data.get("accounts") if isinstance(data, dict) else data
            if isinstance(items, list):
                for a in items:
                    if isinstance(a, dict) and (a.get("username") or a.get("email")) and a.get("password"):
                        out.append(a)
            return out
        except Exception:
            return out

    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(4096)
            f.seek(0)
            first_line = sample.splitlines()[0] if sample.splitlines() else ""
            delim = _guess_delimiter(first_line)

            # Headered CSV
            if delim is not None and any(x in first_line.lower() for x in ("email", "username", "password", "pass")):
                reader = csv.DictReader(f, delimiter=delim)
                for row in reader:
                    if not row:
                        continue
                    norm = {str(k).strip().lower(): (row[k] if k in row else "") for k in row.keys() if k is not None}
                    username = str(norm.get("username") or norm.get("email") or "").strip()
                    password = str(norm.get("password") or norm.get("pass") or "").strip()
                    if not username or not password:
                        continue

                    d = _provider_defaults_for_username(username)
                    name = str(norm.get("name") or "").strip() or username
                    provider = str(norm.get("provider") or d.get("provider") or "Custom").strip() or "Custom"
                    host = str(norm.get("host") or d.get("host") or "").strip()
                    if not host:
                        # Fallback: derive from provider selection
                        host = str(PROVIDERS.get(provider, {}).get("host") or "").strip()
                    port_raw = str(norm.get("port") or d.get("port") or 587).strip()
                    try:
                        port = int(float(port_raw))
                    except Exception:
                        port = int(d.get("port") or 587)

                    def _to_int(v, default: int) -> int:
                        s = str(v).strip().lower()
                        if s in ("1", "true", "yes", "y", "on"):
                            return 1
                        if s in ("0", "false", "no", "n", "off"):
                            return 0
                        return default

                    starttls = _to_int(norm.get("starttls", ""), int(d.get("starttls") or 1))
                    verify_tls = _to_int(norm.get("verify_tls", ""), 1)
                    from_email = str(norm.get("from_email") or "").strip() or username
                    sender_name = str(norm.get("sender_name") or "").strip()
                    ssl_flag = 0
                    try:
                        ssl_flag = int(str(norm.get("ssl") or norm.get("use_ssl") or "").strip() or "0")
                    except Exception:
                        ssl_flag = 0

                    out.append(
                        {
                            "name": name,
                            "provider": provider,
                            "host": host,
                            "port": port,
                            "ssl": 1 if ssl_flag else 0,
                            "starttls": starttls,
                            "verify_tls": verify_tls,
                            "username": username,
                            "password": password,
                            "from_email": from_email,
                            "sender_name": sender_name,
                        }
                    )
                return out

            # Simple list (one account per line)
            for line in f.read().splitlines():
                s = (line or "").strip()
                if not s or s.startswith("#"):
                    continue
                # Try common separators; keep password as-is (can contain punctuation)
                for sep in (":", ";", ",", "\t"):
                    if sep in s:
                        left, right = s.split(sep, 1)
                        username = left.strip()
                        password = right.strip()
                        break
                else:
                    parts = s.split()
                    if len(parts) < 2:
                        continue
                    username, password = parts[0].strip(), " ".join(parts[1:]).strip()

                if not username or not password:
                    continue

                d = _provider_defaults_for_username(username)
                out.append(
                    {
                        "name": username,
                        "provider": d["provider"],
                        "host": str(d.get("host") or "").strip(),
                        "port": int(d["port"]),
                        "starttls": int(d["starttls"]),
                        "verify_tls": 1,
                        "username": username,
                        "password": password,
                        "from_email": username,
                        "sender_name": "",
                        "ssl": 0,
                    }
                )
    except Exception:
        return out

    return out


@dataclass(frozen=True)
class ProxyConfig:
    scheme: str  # socks5 | socks4 | http
    host: str
    port: int
    username: str = ""
    password: str = ""


def parse_proxy_line(line: str) -> Optional[ProxyConfig]:
    s = (line or "").strip()
    if not s or s.startswith("#"):
        return None

    scheme = "socks5"
    rest = s
    if "://" in s:
        scheme, rest = s.split("://", 1)
        scheme = (scheme or "").strip().lower()
        rest = (rest or "").strip()

    user = ""
    pwd = ""
    hostport = rest
    if "@" in rest:
        auth, hostport = rest.split("@", 1)
        if ":" in auth:
            user, pwd = auth.split(":", 1)
        else:
            user = auth

    hostport = hostport.strip()
    if ":" not in hostport:
        return None
    host, port_s = hostport.rsplit(":", 1)
    host = host.strip()
    try:
        port = int(port_s.strip())
    except Exception:
        return None

    if scheme in ("socks5h", "socks5-hostname"):
        scheme = "socks5"
    if scheme not in ("socks5", "socks4", "http"):
        scheme = "socks5"

    if not host or port <= 0:
        return None
    return ProxyConfig(scheme=scheme, host=host, port=port, username=(user or "").strip(), password=(pwd or "").strip())


def read_proxies_from_file(path: str) -> List[ProxyConfig]:
    out: List[ProxyConfig] = []
    if not path:
        return out
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            for line in f.read().splitlines():
                p = parse_proxy_line(line)
                if p is not None:
                    out.append(p)
    except Exception:
        return out

    seen = set()
    uniq: List[ProxyConfig] = []
    for p in out:
        k = (p.scheme, p.host.lower(), int(p.port), p.username, p.password)
        if k in seen:
            continue
        seen.add(k)
        uniq.append(p)
    return uniq


def read_emails_from_csv(path: str) -> List[str]:
    """
    Read emails from a CSV file.

    Notes:
    - We intentionally avoid csv.Sniffer() because it can incorrectly pick a delimiter
      from email addresses (e.g. 'm' in 'mail.ru'), which breaks parsing.
    - Supports: header 'email' OR first column if no header.
    - Supports common delimiters: comma, semicolon, tab, pipe.
    - If no common delimiter is present, treats the file as 1 email per line.
    """
    emails: List[str] = []

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)

        first_line = sample.splitlines()[0] if sample.splitlines() else ""
        common_delims = [",", ";", "\t", "|"]
        delim = next((d for d in common_delims if d in first_line), None)

        if delim is None:
            # One email per line
            for line in f.read().splitlines():
                raw = (line or "").strip()
                if raw:
                    emails.append(raw)
        else:
            rows = list(csv.reader(f, delimiter=delim))
            if not rows:
                return []

            first = [c.strip().lower() for c in rows[0]]
            start_idx = 0
            email_col = 0
            if "email" in first:
                email_col = first.index("email")
                start_idx = 1

            for r in rows[start_idx:]:
                if not r or email_col >= len(r):
                    continue
                raw = (r[email_col] or "").strip()
                if raw:
                    emails.append(raw)

    seen = set()
    out: List[str] = []
    for e in emails:
        k = e.lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(e)
    return out


def normalize_email(raw: str) -> Optional[str]:
    """
    Extract first email-like token from raw cell content.
    Accepts cases like:
      - "user@example.com"
      - "user@example.com; other@example.com"
      - "Name <user@example.com>"
    """
    if not raw:
        return None
    # Remove common invisible/control characters that break parsing.
    s = raw.replace("\x00", "").replace("\u200b", "").replace("\ufeff", "")
    s = s.strip().replace("\n", " ").replace("\r", " ")

    # Split on common separators; try parsing each token.
    for token in re.split(r"[;, \t]+", s):
        t = token.strip().strip("<>").strip()
        if not t:
            continue
        _name, addr = parseaddr(t)
        cand = (addr or t).strip()
        if "@" not in cand:
            continue
        local, _, domain = cand.partition("@")
        if not local or not domain or "." not in domain:
            continue
        return cand
    return None


def build_message(
    from_email: str,
    from_name: str,
    to_email: str,
    subject: str,
    body: str,
    is_html: bool,
    attachments: Iterable[str],
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((from_name, from_email)) if from_name else from_email
    msg["To"] = to_email
    msg["Subject"] = subject

    if is_html:
        msg.set_content("HTML письмо (включите отображение HTML).")
        msg.add_alternative(body or "", subtype="html")
    else:
        msg.set_content(body or "")

    for p in attachments:
        if not p:
            continue
        filename = os.path.basename(p)
        ctype, encoding = mimetypes.guess_type(p)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(p, "rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    return msg


def _default_config() -> dict:
    return {
        "active_account": "",
        "active_template": "",
        "accounts": [],
        "templates": [],
        "rate_per_min": 30,
        "theme": "dark",
    }


def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return _default_config()
        base = _default_config()
        base.update(data)
        # Migration: if sender "name" was stored into from_email (no domain),
        # move it to sender_name and keep username as from_email.
        accounts = base.get("accounts") or []
        changed = False
        for a in accounts:
            if not isinstance(a, dict):
                continue
            user = str(a.get("username", "")).strip()
            fe = str(a.get("from_email", "")).strip()
            sn = str(a.get("sender_name", "")).strip()
            if fe and "@" not in fe and user and "@" in user and not sn:
                a["sender_name"] = fe
                a["from_email"] = user
                changed = True
        if changed:
            base["accounts"] = accounts
            save_config(base)
        return base
    except Exception:
        return _default_config()


def save_config(cfg: dict) -> None:
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)


_BRACE_RE = re.compile(r"\{([^{}]+)\}")


def _apply_inline_brace_variants(text: str) -> str:
    s = text

    def repl(m: re.Match) -> str:
        parts = [p.strip() for p in (m.group(1) or "").split("|")]
        parts = [p for p in parts if p]
        return random.choice(parts) if parts else ""

    for _ in range(50):
        if not _BRACE_RE.search(s):
            break
        s = _BRACE_RE.sub(repl, s)
    return s


def _split_top_level_variants(text: str) -> List[str]:
    parts = [p.strip() for p in text.split("|")]
    return [p for p in parts if p]


def apply_variants(text: str, *, top_index: Optional[int] = None) -> str:
    """
    Variant rules:
    - Inline variants in braces: "{A|B|C}" gets replaced randomly (can appear multiple times)
    - Top-level variants separated by '|':
        - if top_index is None: choose one randomly
        - else: choose by index (modulo number of variants)
    """
    if not text:
        return ""

    s = _apply_inline_brace_variants(text)
    parts = _split_top_level_variants(s)
    if len(parts) >= 2:
        if top_index is None:
            return random.choice(parts)
        return parts[int(top_index) % len(parts)]
    return parts[0] if parts else s


def parse_smtp_codes(text: str) -> set[int]:
    out: set[int] = set()
    if not text:
        return out
    for part in re.split(r"[,\s;]+", str(text)):
        p = (part or "").strip()
        if not p:
            continue
        try:
            out.add(int(p))
        except Exception:
            continue
    return out


def parse_float_relaxed(text: str, default: float) -> float:
    try:
        s = str(text or "").strip()
        if not s:
            return float(default)
        # allow comma decimal separator (common in RU locales)
        s = s.replace(",", ".")
        return float(s)
    except Exception:
        return float(default)


@dataclass(frozen=True)
class TemplateConfig:
    name: str
    subject: str
    body: str
    is_html: bool
    batch_n: int


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    use_ssl: bool
    starttls: bool
    verify_tls: bool
    username: str
    password: str
    from_email: str
    from_name: str = ""


class SenderThread(threading.Thread):
    def __init__(
        self,
        smtp_cfgs: List[SmtpConfig],
        start_smtp_idx: int,
        rotate_every_n: int,
        proxies: List[ProxyConfig],
        use_proxies: bool,
        rotate_proxy_every_n: int,
        rotate_proxy_on_codes: set[int],
        templates: List[TemplateConfig],
        start_template_idx: int,
        rotate_template_every_n: int,
        rotate_template_on_codes: set[int],
        rotate_template_every_s: float,
        rotate_template_fail_streak_n: int,
        emails: List[str],
        autopause_451: bool,
        rotate_on_451: bool,
        rotate_on_codes: set[int],
        pause_451_s: float,
        failure_pause_threshold: int,
        attachments: List[str],
        emails_per_minute: int,
        delay_min_s: float,
        delay_max_s: float,
        stop_event: threading.Event,
        on_log,
        on_progress,
        on_done,
        on_error,
    ) -> None:
        super().__init__(daemon=True)
        self.smtp_cfgs = smtp_cfgs
        self.start_smtp_idx = max(0, int(start_smtp_idx))
        self.rotate_every_n = max(0, int(rotate_every_n))
        self.proxies = proxies
        self.use_proxies = bool(use_proxies)
        self.rotate_proxy_every_n = max(0, int(rotate_proxy_every_n))
        self.rotate_proxy_on_codes = set(int(x) for x in (rotate_proxy_on_codes or set()))
        self.templates = templates
        self.start_template_idx = max(0, int(start_template_idx))
        self.rotate_template_every_n = max(0, int(rotate_template_every_n))
        self.rotate_template_on_codes = set(int(x) for x in (rotate_template_on_codes or set()))
        self.rotate_template_every_s = max(0.0, float(rotate_template_every_s))
        self.rotate_template_fail_streak_n = max(0, int(rotate_template_fail_streak_n))
        self.emails = emails
        self.autopause_451 = bool(autopause_451)
        self.rotate_on_451 = bool(rotate_on_451)
        self.rotate_on_codes = set(int(x) for x in (rotate_on_codes or set()))
        self.pause_451_s = max(0.0, float(pause_451_s))
        self.failure_pause_threshold = max(0, int(failure_pause_threshold))
        self.attachments = attachments
        self.epm = max(1, int(emails_per_minute))
        self.delay_min_s = max(0.0, float(delay_min_s))
        self.delay_max_s = max(0.0, float(delay_max_s))
        self.stop_event = stop_event
        self.on_log = on_log
        self.on_progress = on_progress
        self.on_done = on_done
        self.on_error = on_error

    def _proxy_type(self, scheme: str):
        if socks is None:
            return None
        s = (scheme or "").strip().lower()
        if s == "socks4":
            return socks.SOCKS4
        if s == "http":
            return socks.HTTP
        return socks.SOCKS5

    def _connect(self, cfg: SmtpConfig, proxy: Optional[ProxyConfig]) -> smtplib.SMTP:
        def _tls_context() -> ssl.SSLContext:
            if cfg.verify_tls:
                if certifi is not None:
                    return ssl.create_default_context(cafile=certifi.where())
                return ssl.create_default_context()
            return ssl._create_unverified_context()

        if self.use_proxies and proxy is not None:
            if socks is None:
                raise RuntimeError("PySocks not installed")

            proxy_type = self._proxy_type(proxy.scheme)

            class _ProxySMTP(smtplib.SMTP):
                def _get_socket(self, host, port, timeout):  # type: ignore[override]
                    sck = socks.socksocket()
                    sck.set_proxy(
                        proxy_type,
                        proxy.host,
                        int(proxy.port),
                        username=(proxy.username or None),
                        password=(proxy.password or None),
                    )
                    sck.settimeout(timeout)
                    sck.connect((host, port))
                    return sck

            class _ProxySMTP_SSL(smtplib.SMTP_SSL):
                def _get_socket(self, host, port, timeout):  # type: ignore[override]
                    sck = socks.socksocket()
                    sck.set_proxy(
                        proxy_type,
                        proxy.host,
                        int(proxy.port),
                        username=(proxy.username or None),
                        password=(proxy.password or None),
                    )
                    sck.settimeout(timeout)
                    sck.connect((host, port))
                    ctx = _tls_context()
                    return ctx.wrap_socket(sck, server_hostname=host)

            if cfg.use_ssl:
                smtp = _ProxySMTP_SSL(cfg.host, cfg.port, timeout=30, context=_tls_context())
            else:
                smtp = _ProxySMTP(cfg.host, cfg.port, timeout=30)
        else:
            if cfg.use_ssl:
                smtp = smtplib.SMTP_SSL(cfg.host, cfg.port, timeout=30, context=_tls_context())
            else:
                smtp = smtplib.SMTP(cfg.host, cfg.port, timeout=30)
        smtp.ehlo()
        if (not cfg.use_ssl) and cfg.starttls:
            smtp.starttls(context=_tls_context())
            smtp.ehlo()
        smtp.login(cfg.username, cfg.password)
        return smtp

    def _connect_with_proxy_failover(
        self,
        *,
        cfg: SmtpConfig,
        proxies: List[ProxyConfig],
        proxy_idx: int,
        cur_proxy: Optional[ProxyConfig],
        reason: str,
    ) -> tuple[smtplib.SMTP, int, Optional[ProxyConfig]]:
        if not self.use_proxies or not proxies:
            return self._connect(cfg, None), proxy_idx, None

        if cur_proxy is None:
            proxy_idx = 0
            cur_proxy = proxies[proxy_idx]

        last_err: Optional[Exception] = None
        for attempt in range(len(proxies)):
            p = proxies[proxy_idx]
            try:
                smtp = self._connect(cfg, p)
                if attempt > 0:
                    self.on_log(now_ts(), "-", f"PROXY OK after failover ({reason}) -> {p.scheme}://{p.host}:{p.port}")
                return smtp, proxy_idx, p
            except Exception as e:
                last_err = e
                self.on_log(now_ts(), "-", f"BAD PROXY ({reason}): {p.scheme}://{p.host}:{p.port} ({e})")
                proxy_idx = (proxy_idx + 1) % len(proxies)

        raise RuntimeError(f"All proxies failed ({reason}): {last_err}")

    def run(self) -> None:
        total = len(self.emails)
        sent = 0
        failed = 0
        delay_s = 60.0 / float(self.epm)
        paused = False
        last_idx = 0
        keep_current = False
        failure_streak = 0
        sent_since_rotate = 0
        sent_since_proxy_rotate = 0
        sent_since_template_rotate = 0

        smtp_cfgs = [c for c in (self.smtp_cfgs or []) if isinstance(c, SmtpConfig)]
        if not smtp_cfgs:
            self.on_error("Нет SMTP аккаунтов для отправки.")
            return
        acc_idx = self.start_smtp_idx % len(smtp_cfgs)
        cur_cfg = smtp_cfgs[acc_idx]

        proxies = [p for p in (self.proxies or []) if isinstance(p, ProxyConfig) and p.host and p.port]
        proxy_idx = 0
        cur_proxy: Optional[ProxyConfig] = None
        if self.use_proxies and proxies:
            proxy_idx = 0
            cur_proxy = proxies[proxy_idx]

        templates = [t for t in (self.templates or []) if isinstance(t, TemplateConfig) and t.name]
        if not templates:
            self.on_error("Нет шаблонов для отправки.")
            return
        tpl_idx = self.start_template_idx % len(templates)
        cur_tpl = templates[tpl_idx]
        tpl_fail_streak = 0
        last_tpl_switch_ts = _time.monotonic()

        try:
            if self.use_proxies and proxies:
                smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                    cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="initial connect"
                )
            else:
                smtp = self._connect(cur_cfg, cur_proxy)
        except Exception as e:
            self.on_error(f"SMTP login/connect failed: {e}\n\n{traceback.format_exc()}")
            return

        try:
            for i, to_email in enumerate(self.emails, start=1):
                last_idx = i
                if self.stop_event.is_set():
                    self.on_log(now_ts(), "-", "Stopped by user")
                    paused = True
                    keep_current = True
                    break

                norm_to = normalize_email(to_email)
                if not norm_to:
                    failed += 1
                    failure_streak += 1
                    self.on_log(now_ts(), to_email, "FAILED: invalid email in CSV")
                    remaining = max(0, total - i)
                    self.on_progress(sent, failed, remaining, i, total)
                    continue

                # Time-based template rotation (before building message)
                if self.rotate_template_every_s > 0 and len(templates) >= 2:
                    now = _time.monotonic()
                    if (now - last_tpl_switch_ts) >= float(self.rotate_template_every_s):
                        tpl_idx = (tpl_idx + 1) % len(templates)
                        cur_tpl = templates[tpl_idx]
                        tpl_fail_streak = 0
                        last_tpl_switch_ts = now
                        sent_since_template_rotate = 0
                        self.on_log(now_ts(), "-", f"AUTO-SWITCH TEMPLATE (time) -> {cur_tpl.name}")

                top_idx: Optional[int] = None
                if cur_tpl.batch_n > 0:
                    top_idx = (i - 1) // int(cur_tpl.batch_n)

                subj = apply_variants(cur_tpl.subject, top_index=top_idx)
                body = apply_variants(cur_tpl.body, top_index=top_idx)

                msg = build_message(
                    from_email=cur_cfg.from_email,
                    from_name=cur_cfg.from_name,
                    to_email=norm_to,
                    subject=subj,
                    body=body,
                    is_html=bool(cur_tpl.is_html),
                    attachments=self.attachments,
                )
                send_ok = False
                attempts = 0
                while not send_ok and attempts < 3 and not self.stop_event.is_set():
                    attempts += 1
                    try:
                        smtp.send_message(msg)
                        sent += 1
                        sent_since_rotate += 1
                        sent_since_proxy_rotate += 1
                        sent_since_template_rotate += 1
                        self.on_log(now_ts(), to_email, "SENT")
                        send_ok = True
                        failure_streak = 0
                        tpl_fail_streak = 0
                        break
                    except smtplib.SMTPResponseException as e:
                        code = int(getattr(e, "smtp_code", 0) or 0)
                        msg_txt = str(getattr(e, "smtp_error", b"") or b"")
                        if code and (code in self.rotate_template_on_codes) and len(templates) >= 2:
                            tpl_idx = (tpl_idx + 1) % len(templates)
                            cur_tpl = templates[tpl_idx]
                            sent_since_template_rotate = 0
                            tpl_fail_streak = 0
                            last_tpl_switch_ts = _time.monotonic()
                            self.on_log(now_ts(), to_email, f"SWITCH TEMPLATE due to ({code}) -> {cur_tpl.name} (attempt {attempts}/3)")
                            top_idx = None
                            if cur_tpl.batch_n > 0:
                                top_idx = (i - 1) // int(cur_tpl.batch_n)
                            subj = apply_variants(cur_tpl.subject, top_index=top_idx)
                            body = apply_variants(cur_tpl.body, top_index=top_idx)
                            msg = build_message(
                                from_email=cur_cfg.from_email,
                                from_name=cur_cfg.from_name,
                                to_email=norm_to,
                                subject=subj,
                                body=body,
                                is_html=bool(cur_tpl.is_html),
                                attachments=self.attachments,
                            )
                            continue
                        if (
                            code
                            and (code in self.rotate_proxy_on_codes)
                            and self.use_proxies
                            and len(proxies) >= 2
                        ):
                            proxy_idx = (proxy_idx + 1) % len(proxies)
                            cur_proxy = proxies[proxy_idx]
                            sent_since_proxy_rotate = 0
                            self.on_log(now_ts(), to_email, f"SWITCH PROXY due to ({code}) -> {cur_proxy.scheme}://{cur_proxy.host}:{cur_proxy.port} (attempt {attempts}/3)")
                            try:
                                smtp.quit()
                            except Exception:
                                pass
                            try:
                                smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                    cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason=f"switch due to ({code})"
                                )
                            except Exception as ee:
                                self.on_log(now_ts(), to_email, f"FAILED SWITCH PROXY: {ee}")
                            continue
                        if code and (code in self.rotate_on_codes) and len(smtp_cfgs) >= 2:
                            # Switch account immediately and retry same recipient.
                            acc_idx = (acc_idx + 1) % len(smtp_cfgs)
                            cur_cfg = smtp_cfgs[acc_idx]
                            sent_since_rotate = 0
                            self.on_log(now_ts(), to_email, f"SWITCH due to ({code}) -> {cur_cfg.username} (attempt {attempts}/3)")
                            try:
                                smtp.quit()
                            except Exception:
                                pass
                            try:
                                if self.use_proxies and proxies:
                                    smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                        cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason=f"switch account due to ({code})"
                                    )
                                else:
                                    smtp = self._connect(cur_cfg, cur_proxy)
                            except Exception as ee:
                                self.on_log(now_ts(), to_email, f"FAILED SWITCH ACCOUNT: {ee}")
                            continue
                        if code == 451 and self.rotate_on_451 and len(smtp_cfgs) >= 2:
                            # Switch account immediately and retry same recipient.
                            acc_idx = (acc_idx + 1) % len(smtp_cfgs)
                            cur_cfg = smtp_cfgs[acc_idx]
                            sent_since_rotate = 0
                            self.on_log(now_ts(), to_email, f"SWITCH due to 451 -> {cur_cfg.username} (attempt {attempts}/3)")
                            try:
                                smtp.quit()
                            except Exception:
                                pass
                            try:
                                if self.use_proxies and proxies:
                                    smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                        cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="switch account due to 451"
                                    )
                                else:
                                    smtp = self._connect(cur_cfg, cur_proxy)
                            except Exception as ee:
                                # If switching failed, fall back to the existing behavior below.
                                self.on_log(now_ts(), to_email, f"FAILED SWITCH ACCOUNT: {ee}")
                            continue
                        if self.autopause_451 and code == 451:
                            wait_s = self.pause_451_s or 120.0
                            self.on_log(now_ts(), to_email, f"PAUSE {int(wait_s)}s due to 451 rate limit (attempt {attempts}/3)")
                            # reconnect, wait, retry same recipient
                            try:
                                smtp.quit()
                            except Exception:
                                pass
                            time.sleep(wait_s)
                            try:
                                if self.use_proxies and proxies:
                                    smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                        cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="reconnect after 451 pause"
                                    )
                                else:
                                    smtp = self._connect(cur_cfg, cur_proxy)
                            except Exception:
                                pass
                            continue
                        failed += 1
                        failure_streak += 1
                        tpl_fail_streak += 1
                        self.on_log(now_ts(), to_email, f"FAILED: ({code}) {msg_txt or e}")
                        if self.rotate_template_fail_streak_n > 0 and tpl_fail_streak >= self.rotate_template_fail_streak_n and len(templates) >= 2:
                            tpl_idx = (tpl_idx + 1) % len(templates)
                            cur_tpl = templates[tpl_idx]
                            tpl_fail_streak = 0
                            last_tpl_switch_ts = _time.monotonic()
                            sent_since_template_rotate = 0
                            self.on_log(now_ts(), to_email, f"SWITCH TEMPLATE due to failures -> {cur_tpl.name}")
                        try:
                            smtp.quit()
                        except Exception:
                            pass
                        try:
                            if self.use_proxies and proxies:
                                smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                    cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason=f"reconnect after ({code})"
                                )
                            else:
                                smtp = self._connect(cur_cfg, cur_proxy)
                        except Exception:
                            pass
                        break
                    except Exception as e:
                        failed += 1
                        failure_streak += 1
                        tpl_fail_streak += 1
                        self.on_log(now_ts(), to_email, f"FAILED: {e}")
                        if self.rotate_template_fail_streak_n > 0 and tpl_fail_streak >= self.rotate_template_fail_streak_n and len(templates) >= 2:
                            tpl_idx = (tpl_idx + 1) % len(templates)
                            cur_tpl = templates[tpl_idx]
                            tpl_fail_streak = 0
                            last_tpl_switch_ts = _time.monotonic()
                            sent_since_template_rotate = 0
                            self.on_log(now_ts(), to_email, f"SWITCH TEMPLATE due to failures -> {cur_tpl.name}")
                        try:
                            smtp.quit()
                        except Exception:
                            pass
                        try:
                            if self.use_proxies and proxies:
                                smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                    cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="reconnect after exception"
                                )
                            else:
                                smtp = self._connect(cur_cfg, cur_proxy)
                        except Exception:
                            pass
                        break

                if (not send_ok) and self.failure_pause_threshold and failure_streak >= self.failure_pause_threshold:
                    paused = True
                    keep_current = True  # retry this recipient after user intervention
                    self.on_log(now_ts(), "-", f"AUTO-PAUSE: {failure_streak} failures подряд. Нужна проверка/смена аккаунта.")
                    break

                # Auto-rotate SMTP account after N successful sends
                if self.rotate_every_n > 0 and len(smtp_cfgs) >= 2 and sent_since_rotate >= self.rotate_every_n:
                    sent_since_rotate = 0
                    acc_idx = (acc_idx + 1) % len(smtp_cfgs)
                    cur_cfg = smtp_cfgs[acc_idx]
                    self.on_log(now_ts(), "-", f"AUTO-SWITCH ACCOUNT -> {cur_cfg.username}")
                    try:
                        smtp.quit()
                    except Exception:
                        pass
                    try:
                        if self.use_proxies and proxies:
                            smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                                cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="auto-switch account"
                            )
                        else:
                            smtp = self._connect(cur_cfg, cur_proxy)
                    except Exception as e:
                        failed += 1
                        failure_streak += 1
                        self.on_log(now_ts(), "-", f"FAILED SWITCH ACCOUNT: {e}")

                # Auto-rotate proxy after N successful sends
                if self.rotate_proxy_every_n > 0 and self.use_proxies and len(proxies) >= 2 and sent_since_proxy_rotate >= self.rotate_proxy_every_n:
                    sent_since_proxy_rotate = 0
                    proxy_idx = (proxy_idx + 1) % len(proxies)
                    cur_proxy = proxies[proxy_idx]
                    self.on_log(now_ts(), "-", f"AUTO-SWITCH PROXY -> {cur_proxy.scheme}://{cur_proxy.host}:{cur_proxy.port}")
                    try:
                        smtp.quit()
                    except Exception:
                        pass
                    try:
                        smtp, proxy_idx, cur_proxy = self._connect_with_proxy_failover(
                            cfg=cur_cfg, proxies=proxies, proxy_idx=proxy_idx, cur_proxy=cur_proxy, reason="auto-switch proxy"
                        )
                    except Exception as e:
                        failed += 1
                        failure_streak += 1
                        self.on_log(now_ts(), "-", f"FAILED SWITCH PROXY: {e}")

                # Auto-rotate template after N successful sends
                if self.rotate_template_every_n > 0 and len(templates) >= 2 and sent_since_template_rotate >= self.rotate_template_every_n:
                    sent_since_template_rotate = 0
                    tpl_idx = (tpl_idx + 1) % len(templates)
                    cur_tpl = templates[tpl_idx]
                    tpl_fail_streak = 0
                    last_tpl_switch_ts = _time.monotonic()
                    self.on_log(now_ts(), "-", f"AUTO-SWITCH TEMPLATE -> {cur_tpl.name}")

                remaining = max(0, total - i)
                self.on_progress(sent, failed, remaining, i, total)

                if i < total and not self.stop_event.is_set():
                    if self.delay_max_s > 0 or self.delay_min_s > 0:
                        a = min(self.delay_min_s, self.delay_max_s) if self.delay_max_s else self.delay_min_s
                        b = max(self.delay_min_s, self.delay_max_s) if self.delay_max_s else self.delay_min_s
                        time.sleep(random.uniform(a, b))
                    else:
                        time.sleep(delay_s)

            self.on_done(paused, last_idx, total, keep_current)
        except Exception as e:
            self.on_error(f"{e}\n\n{traceback.format_exc()}")
        finally:
            try:
                smtp.quit()
            except Exception:
                pass


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(f"{APP_NAME} — Tk v{app_version()}")
        self.root.geometry("1050x720")

        self.stop_event = threading.Event()
        self.sender: Optional[SenderThread] = None

        self.cfg = load_config()
        self.theme = StringVar(value=str(self.cfg.get("theme", "dark")))
        self.theme_is_dark = IntVar(value=1 if self.theme.get() == "dark" else 0)
        self._apply_theme()

        self.active_account = StringVar(value=self.cfg.get("active_account", ""))
        self.active_template = StringVar(value=self.cfg.get("active_template", ""))
        self.use_multiple_templates = IntVar(value=int(self.cfg.get("use_multiple_templates", 0)))
        self.rotate_templates = IntVar(value=int(self.cfg.get("rotate_templates", 0)))
        self.rotate_template_every_n = StringVar(value=str(self.cfg.get("rotate_template_every_n", 50)))
        self.rotate_template_every_s = StringVar(value=str(self.cfg.get("rotate_template_every_s", 0)))
        self.rotate_template_on_fail_streak = IntVar(value=int(self.cfg.get("rotate_template_on_fail_streak", 0)))
        self.rotate_template_fail_streak_n = StringVar(value=str(self.cfg.get("rotate_template_fail_streak_n", 3)))
        self.rotate_template_on_codes = IntVar(value=int(self.cfg.get("rotate_template_on_codes", 0)))
        self.rotate_template_codes = StringVar(value=str(self.cfg.get("rotate_template_codes", "451")))

        self.csv_path: Optional[str] = None
        self.emails: List[str] = []
        self.resume_index = 0
        self.rate = StringVar(value=str(self.cfg.get("rate_per_min", 30)))
        self.delay_min_s = StringVar(value=str(self.cfg.get("delay_min_s", 0)))
        self.delay_max_s = StringVar(value=str(self.cfg.get("delay_max_s", 0)))
        self.autopause_451 = IntVar(value=int(self.cfg.get("autopause_451", 1)))
        self.rotate_on_451 = IntVar(value=int(self.cfg.get("rotate_on_451", 0)))
        self.pause_451_s = StringVar(value=str(self.cfg.get("pause_451_s", 120)))
        self.failure_pause_threshold = StringVar(value=str(self.cfg.get("failure_pause_threshold", 5)))
        self.auto_rotate_accounts = IntVar(value=int(self.cfg.get("auto_rotate_accounts", 0)))
        self.rotate_every_n = StringVar(value=str(self.cfg.get("rotate_every_n", 100)))
        self.rotate_account_on_codes = IntVar(value=int(self.cfg.get("rotate_account_on_codes", 0)))
        self.rotate_account_codes = StringVar(value=str(self.cfg.get("rotate_account_codes", "451")))

        # Proxies (SMTP via SOCKS/HTTP CONNECT; requires PySocks)
        self.use_proxies = IntVar(value=int(self.cfg.get("use_proxies", 0)))
        self.rotate_proxies = IntVar(value=int(self.cfg.get("rotate_proxies", 0)))
        self.rotate_proxy_every_n = StringVar(value=str(self.cfg.get("rotate_proxy_every_n", 20)))
        self.rotate_proxy_on_codes = IntVar(value=int(self.cfg.get("rotate_proxy_on_codes", 0)))
        self.rotate_proxy_codes = StringVar(value=str(self.cfg.get("rotate_proxy_codes", "451")))
        self.proxies: List[ProxyConfig] = []
        try:
            for it in (self.cfg.get("proxies") or []):
                if isinstance(it, dict):
                    p = ProxyConfig(
                        scheme=str(it.get("scheme", "socks5") or "socks5"),
                        host=str(it.get("host", "") or "").strip(),
                        port=int(it.get("port", 0) or 0),
                        username=str(it.get("username", "") or ""),
                        password=str(it.get("password", "") or ""),
                    )
                    if p.host and p.port:
                        self.proxies.append(p)
        except Exception:
            self.proxies = []

        # Global sender display name override (applies to all accounts)
        self.sender_name_override = IntVar(value=int(self.cfg.get("sender_name_override", 0)))
        self.sender_name_global = StringVar(value=str(self.cfg.get("sender_name_global", "")))

        # Template editor state
        self.tpl_name = StringVar()
        self.tpl_subject = StringVar()
        self.tpl_is_html = IntVar(value=1)
        self.tpl_batch_n = StringVar(value="0")

        self._build_menu()
        self._build_ui()
        self._bind_clipboard_shortcuts()

    def _apply_theme(self) -> None:
        if sv_ttk is None:
            return
        try:
            sv_ttk.set_theme("dark" if self.theme.get() == "dark" else "light")
        except Exception:
            pass
        # Text/Listbox are not themed by ttk; set basic colors.
        try:
            dark = self.theme.get() == "dark"
            bg = "#1e1e1e" if dark else "#ffffff"
            fg = "#e6e6e6" if dark else "#111111"
            sel_bg = "#2d4f7a" if dark else "#cce3ff"
            sel_fg = "#ffffff" if dark else "#111111"

            for w in (getattr(self, "tpl_body", None), getattr(self, "tpl_list", None), getattr(self, "attach_list", None)):
                if w is None:
                    continue
                try:
                    w.configure(background=bg, foreground=fg)
                except Exception:
                    pass
                try:
                    w.configure(selectbackground=sel_bg, selectforeground=sel_fg)
                except Exception:
                    pass
                try:
                    w.configure(insertbackground=fg)  # caret color (Text)
                except Exception:
                    pass
        except Exception:
            pass

    def toggle_theme(self) -> None:
        self.theme.set("light" if self.theme.get() == "dark" else "dark")
        self.theme_is_dark.set(1 if self.theme.get() == "dark" else 0)
        self._apply_theme()
        self.cfg["theme"] = self.theme.get()
        save_config(self.cfg)

    def set_theme_from_checkbox(self) -> None:
        self.theme.set("dark" if self.theme_is_dark.get() else "light")
        self._apply_theme()
        self.cfg["theme"] = self.theme.get()
        save_config(self.cfg)

    def _bind_clipboard_shortcuts(self) -> None:
        def do_copy(_e=None):
            w = self.root.focus_get()
            if isinstance(w, (Entry, Text)):
                try:
                    w.event_generate("<<Copy>>")
                except Exception:
                    pass
            return "break"

        def do_cut(_e=None):
            w = self.root.focus_get()
            if isinstance(w, (Entry, Text)):
                try:
                    w.event_generate("<<Cut>>")
                except Exception:
                    pass
            return "break"

        def do_paste(_e=None):
            w = self.root.focus_get()
            if isinstance(w, (Entry, Text)):
                try:
                    w.event_generate("<<Paste>>")
                except Exception:
                    try:
                        txt = self.root.clipboard_get()
                        if isinstance(w, Entry):
                            w.insert(END, txt)
                        else:
                            w.insert("insert", txt)
                    except Exception:
                        pass
            return "break"

        # Bind on widget classes (Entry/Text) to avoid double-executing:
        # default widget/class bindings may fire before "bind_all".
        for cls in ("Entry", "Text"):
            for seq in ("<Command-c>", "<Command-C>", "<Control-c>", "<Control-C>"):
                self.root.bind_class(cls, seq, do_copy, add=False)
            for seq in ("<Command-x>", "<Command-X>", "<Control-x>", "<Control-X>"):
                self.root.bind_class(cls, seq, do_cut, add=False)
            for seq in ("<Command-v>", "<Command-V>", "<Control-v>", "<Control-V>"):
                self.root.bind_class(cls, seq, do_paste, add=False)

    def _build_menu(self) -> None:
        m = Menu(self.root)
        filem = Menu(m, tearoff=0)
        filem.add_command(label="Open CSV…", command=self.pick_csv)
        filem.add_separator()
        filem.add_command(label="Quit", command=self.root.destroy)
        m.add_cascade(label="File", menu=filem)

        viewm = Menu(m, tearoff=0)
        viewm.add_command(label="Toggle dark/light", command=self.toggle_theme)
        m.add_cascade(label="View", menu=viewm)

        helpm = Menu(m, tearoff=0)
        helpm.add_command(label="Check updates (Releases)", command=self.open_releases)
        helpm.add_command(label="About", command=self.show_about)
        m.add_cascade(label="Help", menu=helpm)

        self.root.config(menu=m)

    def open_releases(self) -> None:
        if APP_REPO_URL.strip():
            webbrowser.open(APP_REPO_URL.rstrip("/") + "/releases/latest")
        else:
            messagebox.showinfo("Updates", "Set APP_REPO_URL in app_tk.py to enable opening the Releases page.")

    def show_about(self) -> None:
        messagebox.showinfo("About", f"{APP_NAME}\nVersion: {app_version()}")

    def _build_ui(self) -> None:
        top = Frame(self.root)
        top.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=10)

        self.nb = ttk.Notebook(top)
        self.nb.pack(fill=BOTH, expand=True)

        self.tab_accounts = Frame(self.nb)
        self.tab_templates = Frame(self.nb)
        self.tab_send = Frame(self.nb)

        self.nb.add(self.tab_accounts, text="Аккаунты")
        self.nb.add(self.tab_templates, text="Шаблоны")
        self.nb.add(self.tab_send, text="Отправка")

        self._build_accounts_tab()
        self._build_templates_tab()
        self._build_send_tab()

        self._refresh_accounts_tree()
        self._refresh_templates_list()

    def _build_accounts_tab(self) -> None:
        box = ttk.LabelFrame(self.tab_accounts, text="SMTP аккаунты")
        box.pack(fill=BOTH, expand=True, padx=8, pady=8)

        cols = ("name", "provider", "host", "port", "username", "from")
        self.acc_tree = ttk.Treeview(box, columns=cols, show="headings", height=14, selectmode="extended")
        for c, w in zip(cols, (160, 80, 170, 60, 240, 240)):
            self.acc_tree.heading(c, text=c)
            self.acc_tree.column(c, width=w, anchor=W)
        self.acc_tree.pack(side=LEFT, fill=BOTH, expand=True)

        sb = Scrollbar(box, orient=VERTICAL, command=self.acc_tree.yview)
        self.acc_tree.configure(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill="y")

        btns = Frame(self.tab_accounts)
        btns.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Button(btns, text="Добавить", command=self._account_add).pack(side=LEFT)
        ttk.Button(btns, text="Изменить", command=self._account_edit).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Удалить", command=self._account_delete).pack(side=LEFT)
        ttk.Button(btns, text="Удалить выбранные", command=self._account_delete_selected).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Удалить все", command=self._account_delete_all).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Импорт из файла…", command=self._account_import).pack(side=LEFT, padx=12)
        ttk.Button(btns, text="Сделать активным", command=self._account_set_active).pack(side=LEFT, padx=12)
        ttk.Label(btns, text="Активный:").pack(side=LEFT, padx=(20, 6))
        ttk.Label(btns, textvariable=self.active_account, anchor="w").pack(side=LEFT, fill=BOTH, expand=True)

    def _build_templates_tab(self) -> None:
        root = Frame(self.tab_templates)
        root.pack(fill=BOTH, expand=True, padx=8, pady=8)

        left = ttk.LabelFrame(root, text="Шаблоны")
        left.pack(side=LEFT, fill=BOTH, expand=False)
        self.tpl_list = Listbox(left, height=18, width=28)
        self.tpl_list.pack(side=LEFT, fill=BOTH, expand=True, padx=(6, 0), pady=6)
        sb = Scrollbar(left, orient=VERTICAL, command=self.tpl_list.yview)
        self.tpl_list.configure(yscrollcommand=sb.set)
        sb.pack(side=RIGHT, fill="y", pady=6, padx=(0, 6))
        self.tpl_list.bind("<<ListboxSelect>>", lambda _e: self._template_load_selected())

        right = ttk.LabelFrame(root, text="Редактор")
        right.pack(side=LEFT, fill=BOTH, expand=True, padx=(10, 0))

        row1 = Frame(right)
        row1.pack(fill=BOTH, padx=8, pady=6)
        ttk.Label(row1, text="Имя:").pack(side=LEFT)
        ttk.Entry(row1, textvariable=self.tpl_name, width=32).pack(side=LEFT, padx=6)
        ttk.Checkbutton(row1, text="HTML", variable=self.tpl_is_html).pack(side=LEFT, padx=10)
        ttk.Label(row1, text="Пачка N:").pack(side=LEFT, padx=(14, 0))
        ttk.Entry(row1, textvariable=self.tpl_batch_n, width=4).pack(side=LEFT, padx=6)
        ttk.Button(row1, text="Загрузить .html…", command=self._template_load_html_file).pack(side=RIGHT)

        row2 = Frame(right)
        row2.pack(fill=BOTH, padx=8, pady=6)
        ttk.Label(row2, text="Тема:").pack(side=LEFT)
        ttk.Entry(row2, textvariable=self.tpl_subject, width=90).pack(side=LEFT, padx=6, fill=BOTH, expand=True)

        row3 = Frame(right)
        row3.pack(fill=BOTH, padx=8, pady=6, expand=True)
        ttk.Label(row3, text="Тело (HTML/текст):").pack(side=LEFT, anchor="n")
        self.tpl_body = Text(row3, height=18, wrap="word")
        self.tpl_body.pack(side=LEFT, padx=6, fill=BOTH, expand=True)

        row4 = Frame(right)
        row4.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Button(row4, text="Сохранить/обновить", command=self._template_save).pack(side=LEFT)
        ttk.Button(row4, text="Примеры (5)", command=self._template_show_examples).pack(side=LEFT, padx=6)
        ttk.Button(row4, text="Предпросмотр HTML", command=self._template_preview_html).pack(side=LEFT, padx=6)
        ttk.Button(row4, text="Удалить", command=self._template_delete).pack(side=LEFT, padx=6)
        ttk.Button(row4, text="Сделать активным", command=self._template_set_active).pack(side=LEFT, padx=12)
        ttk.Label(row4, text="Активный:").pack(side=LEFT, padx=(20, 6))
        ttk.Label(row4, textvariable=self.active_template, anchor="w").pack(side=LEFT, fill=BOTH, expand=True)

        hint = Message(
            right,
            width=720,
            text="Рандомизация: в теме/теле можно писать варианты через '|' (например 'Привет|Здравствуйте'), "
            "и внутри текста использовать {вариант1|вариант2|вариант3}. На каждого получателя выберется случайно.",
        )
        hint.pack(fill=BOTH, padx=8, pady=(0, 8))

    def _build_send_tab(self) -> None:
        top = Frame(self.tab_send)
        top.pack(side=TOP, fill=BOTH, expand=True, padx=8, pady=8)

        sel = ttk.LabelFrame(top, text="Выбор")
        sel.pack(fill=BOTH)
        r = Frame(sel)
        r.pack(fill=BOTH, padx=8, pady=6)
        ttk.Label(r, text="Аккаунт:").pack(side=LEFT)
        self.send_account_cb = ttk.Combobox(r, textvariable=self.active_account, values=[], state="readonly", width=24)
        self.send_account_cb.pack(side=LEFT, padx=6)
        ttk.Label(r, text="Шаблон:").pack(side=LEFT, padx=(14, 0))
        self.send_template_cb = ttk.Combobox(r, textvariable=self.active_template, values=[], state="readonly", width=24)
        self.send_template_cb.pack(side=LEFT, padx=6)
        ttk.Checkbutton(r, text="несколько", variable=self.use_multiple_templates, command=self._toggle_multi_templates_ui).pack(side=LEFT, padx=8)
        ttk.Button(r, text="Обновить списки", command=self._refresh_send_selectors).pack(side=RIGHT)

        self.multi_tpl_box = ttk.LabelFrame(top, text="Шаблоны (Ctrl/Shift для нескольких)")
        self.multi_tpl_box.pack(fill=BOTH, padx=0, pady=(8, 0))
        self.tpl_send_list = Listbox(self.multi_tpl_box, height=5, selectmode="extended")
        self.tpl_send_list.pack(side=LEFT, fill=BOTH, expand=True, padx=(8, 0), pady=8)
        sb_tpl = Scrollbar(self.multi_tpl_box, orient=VERTICAL, command=self.tpl_send_list.yview)
        self.tpl_send_list.configure(yscrollcommand=sb_tpl.set)
        sb_tpl.pack(side=RIGHT, fill="y", pady=8, padx=(0, 8))

        tpl_rot = Frame(self.multi_tpl_box)
        tpl_rot.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(tpl_rot, text="Авто-смена шаблона каждые N отправок:", variable=self.rotate_templates).pack(side=LEFT)
        ttk.Entry(tpl_rot, textvariable=self.rotate_template_every_n, width=6).pack(side=LEFT, padx=6)
        ttk.Label(tpl_rot, text="(успешных)").pack(side=LEFT, padx=6)
        ttk.Checkbutton(tpl_rot, text="и/или по SMTP кодам:", variable=self.rotate_template_on_codes).pack(side=LEFT, padx=(16, 0))
        ttk.Entry(tpl_rot, textvariable=self.rotate_template_codes, width=16).pack(side=LEFT, padx=6)

        tpl_rot2 = Frame(self.multi_tpl_box)
        tpl_rot2.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Label(tpl_rot2, text="Смена по времени (сек):").pack(side=LEFT)
        ttk.Entry(tpl_rot2, textvariable=self.rotate_template_every_s, width=6).pack(side=LEFT, padx=6)
        ttk.Checkbutton(tpl_rot2, text="Смена при ошибках подряд:", variable=self.rotate_template_on_fail_streak).pack(side=LEFT, padx=(16, 0))
        ttk.Entry(tpl_rot2, textvariable=self.rotate_template_fail_streak_n, width=4).pack(side=LEFT, padx=6)
        ttk.Label(tpl_rot2, text="(шт)").pack(side=LEFT, padx=6)

        r2 = Frame(sel)
        r2.pack(fill=BOTH, padx=8, pady=(0, 6))
        ttk.Checkbutton(r2, text="Единое имя отправителя:", variable=self.sender_name_override).pack(side=LEFT)
        ttk.Entry(r2, textvariable=self.sender_name_global, width=28).pack(side=LEFT, padx=6)
        ttk.Label(r2, text="(например: Магазин)").pack(side=LEFT, padx=6)
        ttk.Checkbutton(r2, text="Тёмная тема", variable=self.theme_is_dark, command=self.set_theme_from_checkbox).pack(side=RIGHT)

        base = ttk.LabelFrame(top, text="База и вложения")
        base.pack(fill=BOTH, pady=(10, 0))
        row3 = Frame(base)
        row3.pack(fill=BOTH, padx=8, pady=6)
        self.csv_label = ttk.Label(row3, text="CSV не выбран", anchor="w")
        self.csv_label.pack(side=LEFT, fill=BOTH, expand=True)
        ttk.Button(row3, text="Выбрать CSV…", command=self.pick_csv).pack(side=RIGHT)

        attach_row = Frame(base)
        attach_row.pack(fill=BOTH, padx=8, pady=6)
        ttk.Label(attach_row, text="Вложения:").pack(side=LEFT, anchor="n")
        self.attach_list = Listbox(attach_row, height=4)
        self.attach_list.pack(side=LEFT, padx=6, fill=BOTH, expand=True)
        btns = Frame(attach_row)
        btns.pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Добавить…", command=self.add_attachment).pack(fill=BOTH, pady=(0, 4))
        ttk.Button(btns, text="Убрать", command=self.remove_attachment).pack(fill=BOTH)

        proxy_row = Frame(base)
        proxy_row.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(proxy_row, text="Использовать прокси", variable=self.use_proxies).pack(side=LEFT)
        ttk.Button(proxy_row, text="Загрузить прокси из файла…", command=self.load_proxies).pack(side=LEFT, padx=8)
        ttk.Button(proxy_row, text="Очистить", command=self.clear_proxies).pack(side=LEFT)
        self.proxy_label = ttk.Label(proxy_row, text=self._proxy_status_text(), anchor="w")
        self.proxy_label.pack(side=LEFT, fill=BOTH, expand=True, padx=8)

        row6 = Frame(base)
        row6.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Label(row6, text="Скорость (писем/мин):").pack(side=LEFT)
        ttk.Entry(row6, textvariable=self.rate, width=6).pack(side=LEFT, padx=6)
        ttk.Label(row6, text="  Задержка (сек) мин/макс:").pack(side=LEFT, padx=(14, 0))
        ttk.Entry(row6, textvariable=self.delay_min_s, width=6).pack(side=LEFT, padx=6)
        ttk.Entry(row6, textvariable=self.delay_max_s, width=6).pack(side=LEFT, padx=6)

        row6b = Frame(base)
        row6b.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(row6b, text="Авто-пауза при 451 (rate limit)", variable=self.autopause_451).pack(side=LEFT)
        ttk.Checkbutton(row6b, text="или переключать аккаунт", variable=self.rotate_on_451).pack(side=LEFT, padx=12)
        ttk.Label(row6b, text="пауза (сек):").pack(side=LEFT, padx=(12, 0))
        ttk.Entry(row6b, textvariable=self.pause_451_s, width=6).pack(side=LEFT, padx=6)
        ttk.Label(row6b, text="  Авто-пауза после ошибок подряд:").pack(side=LEFT, padx=(14, 0))
        ttk.Entry(row6b, textvariable=self.failure_pause_threshold, width=4).pack(side=LEFT, padx=6)

        row6c = Frame(base)
        row6c.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(row6c, text="Авто-смена аккаунта каждые N отправок:", variable=self.auto_rotate_accounts).pack(side=LEFT)
        ttk.Entry(row6c, textvariable=self.rotate_every_n, width=6).pack(side=LEFT, padx=6)
        ttk.Label(row6c, text="(успешных)").pack(side=LEFT, padx=6)

        row6d = Frame(base)
        row6d.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(row6d, text="Переключать аккаунт при SMTP кодах:", variable=self.rotate_account_on_codes).pack(side=LEFT)
        ttk.Entry(row6d, textvariable=self.rotate_account_codes, width=18).pack(side=LEFT, padx=6)
        ttk.Label(row6d, text="например: 451,421,454").pack(side=LEFT, padx=6)

        row6e = Frame(base)
        row6e.pack(fill=BOTH, padx=8, pady=(0, 8))
        ttk.Checkbutton(row6e, text="Авто-смена прокси каждые N отправок:", variable=self.rotate_proxies).pack(side=LEFT)
        ttk.Entry(row6e, textvariable=self.rotate_proxy_every_n, width=6).pack(side=LEFT, padx=6)
        ttk.Label(row6e, text="(успешных)").pack(side=LEFT, padx=6)
        ttk.Checkbutton(row6e, text="и/или по SMTP кодам:", variable=self.rotate_proxy_on_codes).pack(side=LEFT, padx=(16, 0))
        ttk.Entry(row6e, textvariable=self.rotate_proxy_codes, width=16).pack(side=LEFT, padx=6)

        ctrl_box = ttk.LabelFrame(top, text="Управление")
        ctrl_box.pack(side=TOP, fill=BOTH, pady=(10, 0))
        row7 = Frame(ctrl_box)
        row7.pack(fill=BOTH, padx=8, pady=6)
        self.btn_start = ttk.Button(row7, text="Старт", command=self.start)
        self.btn_continue = ttk.Button(row7, text="Продолжить", command=self.continue_sending, state="disabled")
        self.btn_continue_other = ttk.Button(row7, text="Продолжить с другим аккаунтом…", command=self.continue_with_other_account, state="disabled")
        self.btn_next_account = ttk.Button(row7, text="Следующий аккаунт", command=self.next_account)
        self.btn_stop = ttk.Button(row7, text="Стоп", command=self.stop, state="disabled")
        self.btn_start.pack(side=LEFT)
        self.btn_continue.pack(side=LEFT, padx=6)
        self.btn_continue_other.pack(side=LEFT, padx=6)
        self.btn_next_account.pack(side=LEFT, padx=6)
        self.btn_stop.pack(side=LEFT, padx=6)
        ttk.Button(row7, text="Тест на email…", command=self.send_test_dialog).pack(side=LEFT, padx=(14, 6))
        ttk.Button(row7, text="Сохранить лог…", command=self.export_log).pack(side=LEFT)
        self.progress = ttk.Progressbar(row7, orient="horizontal", mode="determinate", maximum=100)
        self.progress.pack(side=LEFT, fill=BOTH, expand=True, padx=10)

        stats_box = ttk.LabelFrame(top, text="Статистика")
        stats_box.pack(side=TOP, fill=BOTH, pady=(10, 0))
        row8 = Frame(stats_box)
        row8.pack(fill=BOTH, padx=8, pady=6)
        self.lbl_total = ttk.Label(row8, text="Всего: 0")
        self.lbl_sent = ttk.Label(row8, text="Отправлено: 0")
        self.lbl_failed = ttk.Label(row8, text="Ошибки: 0")
        self.lbl_remaining = ttk.Label(row8, text="Осталось: 0")
        for w in (self.lbl_total, self.lbl_sent, self.lbl_failed, self.lbl_remaining):
            w.pack(side=LEFT, padx=10)

        log_box = ttk.LabelFrame(top, text="Лог отправки")
        log_box.pack(side=TOP, fill=BOTH, expand=True, pady=(10, 0))
        self.tree = ttk.Treeview(log_box, columns=("time", "email", "status"), show="headings")
        self.tree.heading("time", text="Time")
        self.tree.heading("email", text="Email")
        self.tree.heading("status", text="Status")
        self.tree.column("time", width=160, anchor=W)
        self.tree.column("email", width=320, anchor=W)
        self.tree.column("status", width=520, anchor=W)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        sb2 = Scrollbar(log_box, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side=RIGHT, fill="y")

        self._refresh_send_selectors()
        self._toggle_multi_templates_ui()

    def _toggle_multi_templates_ui(self) -> None:
        try:
            if self.use_multiple_templates.get():
                self.multi_tpl_box.pack(fill=BOTH, padx=0, pady=(8, 0))
                self.send_template_cb.configure(state="disabled")
            else:
                self.multi_tpl_box.pack_forget()
                self.send_template_cb.configure(state="readonly")
        except Exception:
            pass

    def _refresh_send_selectors(self) -> None:
        acc_names = [a.get("name", "") for a in (self.cfg.get("accounts") or []) if a.get("name")]
        tpl_names = [t.get("name", "") for t in (self.cfg.get("templates") or []) if t.get("name")]
        self.send_account_cb["values"] = acc_names
        self.send_template_cb["values"] = tpl_names

        try:
            self.tpl_send_list.delete(0, END)
            for n in tpl_names:
                self.tpl_send_list.insert(END, n)
            # preselect active template in list
            if self.active_template.get() in tpl_names:
                idx = tpl_names.index(self.active_template.get())
                self.tpl_send_list.selection_set(idx)
                self.tpl_send_list.see(idx)
        except Exception:
            pass

        if self.active_account.get() not in acc_names and acc_names:
            self.active_account.set(acc_names[0])
        if self.active_template.get() not in tpl_names and tpl_names:
            self.active_template.set(tpl_names[0])

    def _refresh_accounts_tree(self) -> None:
        for it in self.acc_tree.get_children():
            self.acc_tree.delete(it)
        for acc in self.cfg.get("accounts") or []:
            self.acc_tree.insert(
                "",
                END,
                values=(
                    acc.get("name", ""),
                    acc.get("provider", ""),
                    acc.get("host", ""),
                    acc.get("port", ""),
                    acc.get("username", ""),
                    acc.get("from_email", ""),
                ),
            )
        self._refresh_send_selectors()

    def _refresh_templates_list(self) -> None:
        self.tpl_list.delete(0, END)
        for t in self.cfg.get("templates") or []:
            self.tpl_list.insert(END, t.get("name", ""))
        self._refresh_send_selectors()

    def _selected_account_name(self) -> Optional[str]:
        sel = self.acc_tree.selection()
        if not sel:
            return None
        vals = self.acc_tree.item(sel[0], "values")
        return vals[0] if vals else None

    def _account_add(self) -> None:
        self._account_dialog(None)

    def _account_edit(self) -> None:
        name = self._selected_account_name()
        if not name:
            messagebox.showinfo("Аккаунты", "Выбери аккаунт в списке.")
            return
        acc = next((a for a in (self.cfg.get("accounts") or []) if a.get("name") == name), None)
        if not acc:
            return
        self._account_dialog(acc)

    def _account_delete(self) -> None:
        name = self._selected_account_name()
        if not name:
            messagebox.showinfo("Аккаунты", "Выбери аккаунт в списке.")
            return
        if not messagebox.askyesno("Удалить", f"Удалить аккаунт '{name}'?"):
            return
        self.cfg["accounts"] = [a for a in (self.cfg.get("accounts") or []) if a.get("name") != name]
        if self.active_account.get() == name:
            self.active_account.set("")
            self.cfg["active_account"] = ""
        save_config(self.cfg)
        self._refresh_accounts_tree()

    def _account_delete_selected(self) -> None:
        sel = list(self.acc_tree.selection())
        if not sel:
            messagebox.showinfo("Аккаунты", "Выдели один или несколько аккаунтов в списке.")
            return
        names: List[str] = []
        for iid in sel:
            vals = self.acc_tree.item(iid, "values")
            if vals:
                names.append(str(vals[0]))
        names = [n for n in names if n]
        if not names:
            return
        if not messagebox.askyesno("Удалить", f"Удалить выбранные аккаунты ({len(names)})?"):
            return
        self.cfg["accounts"] = [a for a in (self.cfg.get("accounts") or []) if a.get("name") not in set(names)]
        if self.active_account.get() in set(names):
            self.active_account.set("")
            self.cfg["active_account"] = ""
        save_config(self.cfg)
        self._refresh_accounts_tree()

    def _account_delete_all(self) -> None:
        accounts = self.cfg.get("accounts") or []
        if not accounts:
            return
        if not messagebox.askyesno("Удалить", f"Удалить ВСЕ аккаунты ({len(accounts)})?"):
            return
        self.cfg["accounts"] = []
        self.active_account.set("")
        self.cfg["active_account"] = ""
        save_config(self.cfg)
        self._refresh_accounts_tree()

    def _account_set_active(self) -> None:
        name = self._selected_account_name()
        if not name:
            messagebox.showinfo("Аккаунты", "Выбери аккаунт в списке.")
            return
        self.active_account.set(name)
        self.cfg["active_account"] = name
        save_config(self.cfg)
        self._refresh_send_selectors()

    def _account_import(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[
                ("Accounts (txt/csv)", "*.txt *.csv *.tsv"),
                ("JSON", "*.json"),
                ("All files", "*.*"),
            ]
        )
        if not path:
            return

        items = read_accounts_from_file(path)
        if not items:
            messagebox.showwarning("Импорт", "Не удалось найти аккаунты в файле.\n\nОжидаю формат:\n- email:password (по одной записи на строку)\nили CSV с колонками username/email и password.")
            return

        mode = self._ask_import_mode(total=len(items))
        if mode is None:
            return

        added = 0
        updated = 0
        accounts = self.cfg.get("accounts") or []
        if mode == "replace_all":
            accounts = []
        by_name = {str(a.get("name", "")).strip(): a for a in accounts if isinstance(a, dict) and a.get("name")}

        for raw in items:
            if not isinstance(raw, dict):
                continue

            username = str(raw.get("username") or raw.get("email") or "").strip()
            password = str(raw.get("password") or "").strip()
            if not username or not password:
                continue

            name = str(raw.get("name") or "").strip() or username
            d = _provider_defaults_for_username(username)
            provider = str(raw.get("provider") or d.get("provider") or "Custom").strip() or "Custom"
            host = str(raw.get("host") or d.get("host") or "").strip()
            try:
                port = int(raw.get("port") or d.get("port") or 587)
            except Exception:
                port = int(d.get("port") or 587)

            def _int01(v, default: int) -> int:
                try:
                    s = str(v).strip().lower()
                    if s in ("1", "true", "yes", "y", "on"):
                        return 1
                    if s in ("0", "false", "no", "n", "off"):
                        return 0
                except Exception:
                    pass
                return default

            starttls = _int01(raw.get("starttls", None), int(d.get("starttls") or 1))
            verify_tls = _int01(raw.get("verify_tls", None), 1)
            from_email = str(raw.get("from_email") or "").strip() or username
            sender_name = str(raw.get("sender_name") or "").strip()

            new_acc = {
                "name": name,
                "provider": provider,
                "host": host,
                "port": port,
                "ssl": int(_int01(raw.get("ssl", raw.get("use_ssl", 0)), 0)),
                "starttls": int(starttls),
                "verify_tls": int(verify_tls),
                "username": username,
                "password": password,
                "from_email": from_email,
                "sender_name": sender_name,
            }

            if name in by_name:
                # Update existing by name
                if mode == "add_only":
                    continue
                by_name[name].update(new_acc)
                updated += 1
            else:
                accounts.append(new_acc)
                by_name[name] = new_acc
                added += 1

        self.cfg["accounts"] = accounts
        if not self.active_account.get().strip() and accounts:
            self.active_account.set(accounts[0].get("name", ""))
            self.cfg["active_account"] = self.active_account.get()

        save_config(self.cfg)
        self._refresh_accounts_tree()
        messagebox.showinfo("Импорт", f"Готово.\nДобавлено: {added}\nОбновлено: {updated}\n\nВажно: пароли сохраняются в config.json в открытом виде.")

    def _ask_import_mode(self, *, total: int) -> Optional[str]:
        """
        Returns one of: 'add_update' | 'add_only' | 'replace_all' | None (cancel)
        """
        win = Toplevel(self.root)
        win.title("Импорт аккаунтов")
        win.geometry("520x220")
        win.transient(self.root)
        try:
            win.grab_set()
        except Exception:
            pass

        v = StringVar(value="add_update")

        frm = ttk.Frame(win)
        frm.pack(fill=BOTH, expand=True, padx=12, pady=12)

        ttk.Label(frm, text=f"Найдено аккаунтов в файле: {total}").pack(anchor="w", pady=(0, 8))
        ttk.Label(frm, text="Выбери, как применить импорт:").pack(anchor="w", pady=(0, 8))

        ttk.Radiobutton(frm, text="Добавить новые + обновить существующие (по имени)", variable=v, value="add_update").pack(anchor="w")
        ttk.Radiobutton(frm, text="Только добавить новые (не трогать существующие)", variable=v, value="add_only").pack(anchor="w", pady=(4, 0))
        ttk.Radiobutton(frm, text="Заменить ВСЕ аккаунты списком из файла", variable=v, value="replace_all").pack(anchor="w", pady=(4, 0))

        out: dict = {"mode": None}

        def ok():
            out["mode"] = v.get()
            win.destroy()

        def cancel():
            out["mode"] = None
            win.destroy()

        btns = ttk.Frame(frm)
        btns.pack(fill=BOTH, pady=(14, 0))
        ttk.Button(btns, text="Импортировать", command=ok).pack(side=LEFT)
        ttk.Button(btns, text="Отмена", command=cancel).pack(side=LEFT, padx=8)

        win.protocol("WM_DELETE_WINDOW", cancel)
        self.root.wait_window(win)
        return out["mode"]

    def _account_dialog(self, acc: Optional[dict]) -> None:
        win = Toplevel(self.root)
        win.title("Аккаунт SMTP")
        win.geometry("760x260")

        v_name = StringVar(value=(acc.get("name") if acc else ""))
        v_provider = StringVar(value=(acc.get("provider") if acc else "Mail.ru"))
        v_host = StringVar(value=(acc.get("host") if acc else PROVIDERS.get(v_provider.get(), PROVIDERS["Custom"])["host"]))
        v_port = StringVar(value=str(acc.get("port") if acc else PROVIDERS.get(v_provider.get(), PROVIDERS["Custom"])["port"]))
        v_ssl = IntVar(value=int(acc.get("ssl") if acc else 0))
        v_starttls = IntVar(value=int(acc.get("starttls") if acc else 1))
        v_verify = IntVar(value=int(acc.get("verify_tls") if acc else 1))
        v_user = StringVar(value=(acc.get("username") if acc else ""))
        v_pass = StringVar(value=(acc.get("password") if acc else ""))
        v_from = StringVar(value=(acc.get("from_email") if acc else ""))
        v_sender = StringVar(value=(acc.get("sender_name") if acc else ""))

        def apply_defaults(_e=None):
            p = v_provider.get()
            cfg = PROVIDERS.get(p, PROVIDERS["Custom"])
            if acc is None:
                v_host.set(cfg["host"])
                v_port.set(str(465 if v_ssl.get() else cfg["port"]))
                v_starttls.set(0 if v_ssl.get() else int(cfg["starttls"]))

        def toggle_ssl():
            # SSL (465) and STARTTLS are mutually exclusive.
            if v_ssl.get():
                v_starttls.set(0)
                if (v_port.get() or "").strip() == "587":
                    v_port.set("465")
            else:
                if (v_port.get() or "").strip() == "465":
                    v_port.set("587")

        def toggle_starttls():
            if v_starttls.get():
                v_ssl.set(0)
                if (v_port.get() or "").strip() == "465":
                    v_port.set("587")

        frm = Frame(win)
        frm.pack(fill=BOTH, expand=True, padx=10, pady=10)

        r1 = Frame(frm)
        r1.pack(fill=BOTH, pady=4)
        Label(r1, text="Имя:").pack(side=LEFT)
        Entry(r1, textvariable=v_name, width=22).pack(side=LEFT, padx=6)
        Label(r1, text="Провайдер:").pack(side=LEFT, padx=(14, 0))
        cb = ttk.Combobox(r1, textvariable=v_provider, values=list(PROVIDERS.keys()), width=12, state="readonly")
        cb.pack(side=LEFT, padx=6)
        cb.bind("<<ComboboxSelected>>", apply_defaults)
        Label(r1, text="host:").pack(side=LEFT, padx=(14, 0))
        Entry(r1, textvariable=v_host, width=22).pack(side=LEFT, padx=6)
        Label(r1, text="port:").pack(side=LEFT, padx=(14, 0))
        Entry(r1, textvariable=v_port, width=6).pack(side=LEFT, padx=6)

        r2 = Frame(frm)
        r2.pack(fill=BOTH, pady=4)
        Checkbutton(r2, text="SSL (465)", variable=v_ssl, command=toggle_ssl).pack(side=LEFT)
        Checkbutton(r2, text="STARTTLS (587)", variable=v_starttls, command=toggle_starttls).pack(side=LEFT, padx=12)
        Checkbutton(r2, text="Проверять TLS", variable=v_verify).pack(side=LEFT, padx=12)

        r3 = Frame(frm)
        r3.pack(fill=BOTH, pady=4)
        Label(r3, text="Логин (email):").pack(side=LEFT)
        Entry(r3, textvariable=v_user, width=30).pack(side=LEFT, padx=6)
        Label(r3, text="Пароль / app-password:").pack(side=LEFT, padx=(14, 0))
        Entry(r3, textvariable=v_pass, width=24, show="•").pack(side=LEFT, padx=6)

        r4 = Frame(frm)
        r4.pack(fill=BOTH, pady=4)
        Label(r4, text="From email (если отличается):").pack(side=LEFT)
        Entry(r4, textvariable=v_from, width=34).pack(side=LEFT, padx=6)
        Label(r4, text="Имя отправителя:").pack(side=LEFT, padx=(14, 0))
        Entry(r4, textvariable=v_sender, width=22).pack(side=LEFT, padx=6)

        def on_save():
            name = v_name.get().strip()
            if not name:
                messagebox.showwarning("Аккаунт", "Укажи имя аккаунта.")
                return
            host = v_host.get().strip()
            if not host:
                messagebox.showwarning("Аккаунт", "Укажи SMTP host.")
                return
            try:
                port = int(v_port.get())
            except Exception:
                messagebox.showwarning("Аккаунт", "Port должен быть числом.")
                return
            user = v_user.get().strip()
            if "@" not in user:
                messagebox.showwarning("Аккаунт", "Логин должен быть email.")
                return
            pwd = v_pass.get()
            if not pwd:
                messagebox.showwarning("Аккаунт", "Укажи пароль/app-password.")
                return
            from_email = v_from.get().strip() or user
            sender_name = v_sender.get().strip()

            # If user pasted a non-email into from_email (e.g. "info"), treat it as sender_name.
            if from_email and "@" not in from_email:
                if not sender_name:
                    sender_name = from_email
                from_email = user

            if "@" not in from_email:
                messagebox.showwarning("Аккаунт", "From email должен быть email адресом.")
                return

            new_acc = {
                "name": name,
                "provider": v_provider.get(),
                "host": host,
                "port": port,
                "ssl": int(v_ssl.get()),
                "starttls": int(v_starttls.get()),
                "verify_tls": int(v_verify.get()),
                "username": user,
                "password": pwd,
                "from_email": from_email,
                "sender_name": sender_name,
            }

            accounts = self.cfg.get("accounts") or []
            accounts = [a for a in accounts if a.get("name") != name]
            accounts.append(new_acc)
            self.cfg["accounts"] = accounts
            if not self.active_account.get():
                self.active_account.set(name)
                self.cfg["active_account"] = name
            save_config(self.cfg)
            self._refresh_accounts_tree()
            win.destroy()

        r5 = Frame(frm)
        r5.pack(fill=BOTH, pady=(12, 0))
        Button(r5, text="Сохранить", command=on_save).pack(side=LEFT)
        Button(r5, text="Отмена", command=win.destroy).pack(side=LEFT, padx=6)

    def _template_load_selected(self) -> None:
        sel = self.tpl_list.curselection()
        if not sel:
            return
        name = self.tpl_list.get(sel[0])
        t = next((x for x in (self.cfg.get("templates") or []) if x.get("name") == name), None)
        if not t:
            return
        self.tpl_name.set(t.get("name", ""))
        self.tpl_subject.set(t.get("subject", ""))
        self.tpl_is_html.set(int(bool(t.get("is_html", True))))
        self.tpl_batch_n.set(str(t.get("batch_n", 0)))
        self.tpl_body.delete("1.0", END)
        self.tpl_body.insert("1.0", t.get("body", ""))

    def _template_load_html_file(self) -> None:
        try:
            # Tk on macOS expects patterns separated by spaces, not semicolons.
            path = filedialog.askopenfilename(filetypes=[("HTML", "*.html *.htm"), ("All files", "*.*")])
            if not path:
                return
            try:
                with open(path, "r", encoding="utf-8") as f:
                    html = f.read()
            except Exception:
                with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
                    html = f.read()
            self.tpl_is_html.set(1)
            self.tpl_body.delete("1.0", END)
            self.tpl_body.insert("1.0", html)
        except Exception as e:
            messagebox.showerror("HTML", f"Не удалось загрузить HTML: {e}\n\n{traceback.format_exc()}")

    def _template_save(self) -> None:
        name = self.tpl_name.get().strip()
        if not name:
            messagebox.showwarning("Шаблон", "Укажи имя шаблона.")
            return
        subj = self.tpl_subject.get().strip()
        if not subj:
            messagebox.showwarning("Шаблон", "Укажи тему шаблона.")
            return
        try:
            batch_n = int(self.tpl_batch_n.get() or "0")
        except Exception:
            messagebox.showwarning("Шаблон", "Пачка N должна быть числом (0 = случайно).")
            return
        if batch_n < 0:
            batch_n = 0
        body = self.tpl_body.get("1.0", END).rstrip("\n")
        if not body:
            messagebox.showwarning("Шаблон", "Укажи тело шаблона.")
            return
        t = {"name": name, "subject": subj, "body": body, "is_html": int(bool(self.tpl_is_html.get())), "batch_n": batch_n}
        templates = self.cfg.get("templates") or []
        templates = [x for x in templates if x.get("name") != name]
        templates.append(t)
        self.cfg["templates"] = templates
        if not self.active_template.get():
            self.active_template.set(name)
            self.cfg["active_template"] = name
        save_config(self.cfg)
        self._refresh_templates_list()

    def _template_delete(self) -> None:
        name = self.tpl_name.get().strip()
        if not name:
            messagebox.showinfo("Шаблоны", "Выбери шаблон или введи имя.")
            return
        if not messagebox.askyesno("Удалить", f"Удалить шаблон '{name}'?"):
            return
        self.cfg["templates"] = [t for t in (self.cfg.get("templates") or []) if t.get("name") != name]
        if self.active_template.get() == name:
            self.active_template.set("")
            self.cfg["active_template"] = ""
        save_config(self.cfg)
        self._refresh_templates_list()

    def _template_set_active(self) -> None:
        name = self.tpl_name.get().strip()
        if not name:
            messagebox.showwarning("Шаблон", "Укажи/выбери имя шаблона.")
            return
        names = [t.get("name") for t in (self.cfg.get("templates") or [])]
        if name not in names:
            messagebox.showwarning("Шаблон", "Сначала сохрани шаблон.")
            return
        self.active_template.set(name)
        self.cfg["active_template"] = name
        save_config(self.cfg)
        self._refresh_send_selectors()

    def _template_show_examples(self) -> None:
        subj = self.tpl_subject.get().strip()
        body = self.tpl_body.get("1.0", END).rstrip("\n")
        if not subj and not body:
            messagebox.showinfo("Примеры", "Сначала заполни тему/тело шаблона.")
            return
        try:
            batch_n = int(self.tpl_batch_n.get() or "0")
        except Exception:
            batch_n = 0
        lines = []
        for i in range(5):
            top_idx = (i // batch_n) if batch_n > 0 else None
            lines.append(f"--- Example {i+1} ---")
            if subj:
                lines.append("Subject: " + apply_variants(subj, top_index=top_idx))
            if body:
                preview = apply_variants(body, top_index=top_idx)
                preview = preview[:800] + ("\n... (truncated)" if len(preview) > 800 else "")
                lines.append(preview)
            lines.append("")
        messagebox.showinfo("Примеры", "\n".join(lines))

    def _template_preview_html(self) -> None:
        # Write a temporary HTML file and open in default browser.
        is_html = bool(self.tpl_is_html.get())
        if not is_html:
            messagebox.showinfo("Предпросмотр", "Включи HTML в шаблоне, чтобы сделать предпросмотр.")
            return
        body = self.tpl_body.get("1.0", END).rstrip("\n")
        if not body.strip():
            messagebox.showinfo("Предпросмотр", "Шаблон пустой.")
            return
        try:
            # Apply randomization once for preview
            html = apply_variants(body)
            path = os.path.join(APP_DIR, "_preview.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            webbrowser.open(f"file://{path}")
        except Exception as e:
            messagebox.showerror("Предпросмотр", f"Не удалось открыть предпросмотр: {e}")

    def pick_csv(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            emails = read_emails_from_csv(path)
        except Exception as e:
            messagebox.showerror("Ошибка CSV", str(e))
            return
        self.csv_path = path
        self.emails = emails
        self.resume_index = 0
        self.csv_label.config(text=f"{os.path.basename(path)} (email: {len(emails)})")
        self.lbl_total.config(text=f"Всего: {len(emails)}")
        self.lbl_sent.config(text="Отправлено: 0")
        self.lbl_failed.config(text="Ошибки: 0")
        self.lbl_remaining.config(text=f"Осталось: {len(emails)}")
        self.progress["value"] = 0

    def add_attachment(self) -> None:
        paths = filedialog.askopenfilenames()
        for p in paths:
            if p:
                self.attach_list.insert(END, p)

    def remove_attachment(self) -> None:
        sel = list(self.attach_list.curselection())
        for idx in reversed(sel):
            self.attach_list.delete(idx)

    def _get_attachments(self) -> List[str]:
        return [self.attach_list.get(i) for i in range(self.attach_list.size())]

    def _validate(self) -> Optional[str]:
        if not self.csv_path or not self.emails:
            return "Выбери CSV с email адресами."
        if not self.active_account.get().strip():
            return "Выбери активный SMTP аккаунт (вкладка Аккаунты)."
        if not self.active_template.get().strip():
            return "Выбери активный шаблон (вкладка Шаблоны)."
        if self.use_multiple_templates.get():
            try:
                if not list(self.tpl_send_list.curselection()):
                    return "Выбери один или несколько шаблонов (Ctrl/Shift)."
            except Exception:
                return "Выбери один или несколько шаблонов (Ctrl/Shift)."
        try:
            int(self.rate.get())
        except Exception:
            return "Скорость должна быть числом."
        if self.auto_rotate_accounts.get():
            try:
                n = int(self.rotate_every_n.get() or "0")
            except Exception:
                return "N для авто-смены аккаунта должно быть числом."
            if n <= 0:
                return "N для авто-смены аккаунта должно быть > 0."
        if self.rotate_account_on_codes.get():
            codes = parse_smtp_codes(self.rotate_account_codes.get())
            if not codes:
                return "Укажи SMTP-коды для переключения аккаунта (например: 451,421)."
        if self.use_proxies.get():
            if socks is None:
                return "Для прокси нужен пакет PySocks. Установи зависимости заново (pip install -r requirements.txt)."
            if not self.proxies:
                return "Прокси включены, но список пуст. Нажми «Загрузить прокси из файла…»."
            if self.rotate_proxies.get():
                try:
                    n = int(self.rotate_proxy_every_n.get() or "0")
                except Exception:
                    return "N для авто-смены прокси должно быть числом."
                if n <= 0:
                    return "N для авто-смены прокси должно быть > 0."
            if self.rotate_proxy_on_codes.get():
                codes = parse_smtp_codes(self.rotate_proxy_codes.get())
                if not codes:
                    return "Укажи SMTP-коды для переключения прокси (например: 451,421)."
        if self.use_multiple_templates.get():
            if self.rotate_templates.get():
                try:
                    n = int(self.rotate_template_every_n.get() or "0")
                except Exception:
                    return "N для авто-смены шаблона должно быть числом."
                if n <= 0:
                    return "N для авто-смены шаблона должно быть > 0."
            # time-based (allow float seconds)
            tsec = parse_float_relaxed(self.rotate_template_every_s.get(), 0.0)
            if tsec < 0:
                return "Время для авто-смены шаблона должно быть >= 0."
            if self.rotate_template_on_fail_streak.get():
                try:
                    n = int(self.rotate_template_fail_streak_n.get() or "0")
                except Exception:
                    return "Ошибки подряд для смены шаблона должно быть числом."
                if n <= 0:
                    return "Ошибки подряд для смены шаблона должно быть > 0."
            if self.rotate_template_on_codes.get():
                codes = parse_smtp_codes(self.rotate_template_codes.get())
                if not codes:
                    return "Укажи SMTP-коды для переключения шаблона (например: 451,421)."
        return None

    def _proxy_status_text(self) -> str:
        n = len(self.proxies or [])
        on = bool(self.use_proxies.get())
        return f"Прокси: {n} ({'ON' if on else 'OFF'})"

    def load_proxies(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[
                ("Proxies (txt)", "*.txt"),
                ("All files", "*.*"),
            ]
        )
        if not path:
            return
        items = read_proxies_from_file(path)
        if not items:
            messagebox.showwarning(
                "Прокси",
                "Не удалось прочитать прокси.\n\nФормат строк:\n- host:port\n- user:pass@host:port\n- socks5://user:pass@host:port\n- http://host:port",
            )
            return
        self.proxies = items
        self.cfg["proxies"] = [
            {"scheme": p.scheme, "host": p.host, "port": p.port, "username": p.username, "password": p.password} for p in self.proxies
        ]
        save_config(self.cfg)
        try:
            self.proxy_label.config(text=self._proxy_status_text())
        except Exception:
            pass
        messagebox.showinfo("Прокси", f"Загружено прокси: {len(self.proxies)}")

    def clear_proxies(self) -> None:
        if not self.proxies:
            return
        if not messagebox.askyesno("Прокси", "Удалить все загруженные прокси?"):
            return
        self.proxies = []
        self.cfg["proxies"] = []
        # Optionally disable proxies as well
        self.use_proxies.set(0)
        self.cfg["use_proxies"] = 0
        save_config(self.cfg)
        try:
            self.proxy_label.config(text=self._proxy_status_text())
        except Exception:
            pass

    def _set_running(self, running: bool) -> None:
        self.btn_start.config(state=("disabled" if running else "normal"))
        self.btn_stop.config(state=("normal" if running else "disabled"))
        self.btn_continue.config(state=("disabled" if running else ("normal" if self.resume_index > 0 else "disabled")))
        self.btn_continue_other.config(state=("disabled" if running else ("normal" if self.resume_index > 0 else "disabled")))
        self.btn_next_account.config(state=("disabled" if running else "normal"))

    def _ui_log(self, ts: str, email: str, status: str) -> None:
        self.tree.insert("", END, values=(ts, email, status))
        self.tree.yview_moveto(1.0)

    def _ui_progress(self, sent: int, failed: int, remaining: int, done: int, total: int) -> None:
        self.lbl_sent.config(text=f"Отправлено: {sent}")
        self.lbl_failed.config(text=f"Ошибки: {failed}")
        self.lbl_remaining.config(text=f"Осталось: {remaining}")
        self.progress["value"] = 0 if total <= 0 else int((done / total) * 100)

    def start(self) -> None:
        self.resume_index = 0
        self._run_send(reset_log=True)

    def _run_send(self, *, reset_log: bool) -> None:
        err = self._validate()
        if err:
            messagebox.showwarning("Проверка", err)
            return

        if reset_log:
            self.tree.delete(*self.tree.get_children())
            self.progress["value"] = 0
        self.stop_event.clear()
        self._set_running(True)

        acc_name = self.active_account.get().strip()
        tpl_name = self.active_template.get().strip()
        acc = next((a for a in (self.cfg.get("accounts") or []) if a.get("name") == acc_name), None)
        tpl = next((t for t in (self.cfg.get("templates") or []) if t.get("name") == tpl_name), None)
        if not acc or not tpl:
            messagebox.showerror("Ошибка", "Аккаунт или шаблон не найдены. Нажми «Обновить списки».")
            self._set_running(False)
            return

        smtp_cfg = SmtpConfig(
            host=str(acc.get("host", "")).strip(),
            port=int(acc.get("port", 587)),
            use_ssl=bool(int(acc.get("ssl", 0) or 0)),
            starttls=bool(int(acc.get("starttls", 1))),
            verify_tls=bool(int(acc.get("verify_tls", 1))),
            username=str(acc.get("username", "")).strip(),
            password=str(acc.get("password", "")),
            from_email=str(acc.get("from_email", "")).strip() or str(acc.get("username", "")).strip(),
            from_name=(self.sender_name_global.get().strip() if self.sender_name_override.get() else str(acc.get("sender_name", "")).strip()),
        )
        accounts_all = [a for a in (self.cfg.get("accounts") or []) if isinstance(a, dict) and a.get("name")]
        smtp_cfgs: List[SmtpConfig] = []
        for a in accounts_all:
            smtp_cfgs.append(
                SmtpConfig(
                    host=str(a.get("host", "")).strip(),
                    port=int(a.get("port", 587)),
                    use_ssl=bool(int(a.get("ssl", 0) or 0)),
                    starttls=bool(int(a.get("starttls", 1))),
                    verify_tls=bool(int(a.get("verify_tls", 1))),
                    username=str(a.get("username", "")).strip(),
                    password=str(a.get("password", "")),
                    from_email=str(a.get("from_email", "")).strip() or str(a.get("username", "")).strip(),
                    from_name=(self.sender_name_global.get().strip() if self.sender_name_override.get() else str(a.get("sender_name", "")).strip()),
                )
            )
        if not smtp_cfgs:
            smtp_cfgs = [smtp_cfg]

        # Start from currently active account within the list
        start_idx = 0
        try:
            start_idx = next(i for i, a in enumerate(accounts_all) if str(a.get("name", "")).strip() == acc_name)
        except Exception:
            start_idx = 0

        try:
            rotate_n = int(self.rotate_every_n.get() or "0")
        except Exception:
            rotate_n = 0
        rotate_n = rotate_n if self.auto_rotate_accounts.get() else 0
        rotate_codes = parse_smtp_codes(self.rotate_account_codes.get()) if self.rotate_account_on_codes.get() else set()

        try:
            dmin = parse_float_relaxed(self.delay_min_s.get(), 0.0)
            dmax = parse_float_relaxed(self.delay_max_s.get(), 0.0)
        except Exception:
            messagebox.showwarning("Проверка", "Задержка мин/макс должна быть числом (сек).")
            self._set_running(False)
            return

        try:
            pause451 = parse_float_relaxed(self.pause_451_s.get(), 120.0)
        except Exception:
            messagebox.showwarning("Проверка", "Пауза при 451 должна быть числом (сек).")
            self._set_running(False)
            return

        try:
            fail_pause_n = int(self.failure_pause_threshold.get() or "0")
        except Exception:
            messagebox.showwarning("Проверка", "Авто-пауза после ошибок подряд должна быть числом.")
            self._set_running(False)
            return

        def on_log(ts: str, email: str, status: str) -> None:
            self.root.after(0, lambda: self._ui_log(ts, email, status))

        def on_progress(sent: int, failed: int, remaining: int, done: int, total: int) -> None:
            self.root.after(0, lambda: self._ui_progress(sent, failed, remaining, done, total))

        def on_done(paused: bool, last_idx: int, total: int, keep_current: bool) -> None:
            def _finish():
                if paused:
                    # last_idx is 1-based within the slice; if keep_current, retry current recipient
                    adv = max(0, last_idx - (1 if keep_current else 0))
                    self.resume_index = min(len(self.emails), self.resume_index + adv)
                    self._ui_log(now_ts(), "-", f"PAUSED (next index: {self.resume_index + 1})")
                else:
                    self.resume_index = 0
                    self._ui_log(now_ts(), "-", "DONE")
                self._set_running(False)

            self.root.after(0, _finish)

        def on_error(msg: str) -> None:
            self.root.after(0, lambda: (self._ui_log(now_ts(), "-", f"ERROR: {msg}"), messagebox.showerror("Ошибка", msg), self._set_running(False)))

        send_list = self.emails[self.resume_index :]

        # Build template rotation set
        templates_all = [t for t in (self.cfg.get("templates") or []) if isinstance(t, dict) and t.get("name")]
        tpl_by_name = {str(t.get("name")): t for t in templates_all}
        selected_tpl_names: List[str] = []
        if self.use_multiple_templates.get():
            try:
                idxs = list(self.tpl_send_list.curselection())
                names_all = [x.get("name", "") for x in templates_all]
                for idx in idxs:
                    if 0 <= idx < len(names_all):
                        selected_tpl_names.append(str(names_all[idx]))
            except Exception:
                selected_tpl_names = []
        if not selected_tpl_names:
            selected_tpl_names = [tpl_name]

        tpl_cfgs: List[TemplateConfig] = []
        for nm in selected_tpl_names:
            t = tpl_by_name.get(nm)
            if not t:
                continue
            tpl_cfgs.append(
                TemplateConfig(
                    name=str(t.get("name", "")),
                    subject=str(t.get("subject", "")).strip(),
                    body=str(t.get("body", "")),
                    is_html=bool(int(t.get("is_html", 1) or 0)),
                    batch_n=int(t.get("batch_n", 0) or 0),
                )
            )
        if not tpl_cfgs and tpl:
            tpl_cfgs = [
                TemplateConfig(
                    name=str(tpl.get("name", "")) or "template",
                    subject=str(tpl.get("subject", "")).strip(),
                    body=str(tpl.get("body", "")),
                    is_html=bool(int(tpl.get("is_html", 1) or 0)),
                    batch_n=int(tpl.get("batch_n", 0) or 0),
                )
            ]

        start_tpl_idx = 0
        if self.use_multiple_templates.get() and tpl_cfgs:
            # start from first selected
            start_tpl_idx = 0
        rotate_tpl_n = int(self.rotate_template_every_n.get() or "0") if (self.use_multiple_templates.get() and self.rotate_templates.get()) else 0
        rotate_tpl_codes = parse_smtp_codes(self.rotate_template_codes.get()) if (self.use_multiple_templates.get() and self.rotate_template_on_codes.get()) else set()
        rotate_tpl_s = parse_float_relaxed(self.rotate_template_every_s.get(), 0.0) if self.use_multiple_templates.get() else 0.0
        rotate_tpl_fail_n = int(self.rotate_template_fail_streak_n.get() or "0") if (self.use_multiple_templates.get() and self.rotate_template_on_fail_streak.get()) else 0

        self.sender = SenderThread(
            smtp_cfgs=smtp_cfgs,
            start_smtp_idx=start_idx,
            rotate_every_n=rotate_n,
            proxies=list(self.proxies or []),
            use_proxies=bool(self.use_proxies.get()),
            rotate_proxy_every_n=(int(self.rotate_proxy_every_n.get() or "0") if self.rotate_proxies.get() else 0),
            rotate_proxy_on_codes=(parse_smtp_codes(self.rotate_proxy_codes.get()) if self.rotate_proxy_on_codes.get() else set()),
            templates=tpl_cfgs,
            start_template_idx=start_tpl_idx,
            rotate_template_every_n=rotate_tpl_n,
            rotate_template_on_codes=rotate_tpl_codes,
            rotate_template_every_s=rotate_tpl_s,
            rotate_template_fail_streak_n=rotate_tpl_fail_n,
            emails=send_list,
            autopause_451=bool(self.autopause_451.get()),
            rotate_on_451=bool(self.rotate_on_451.get()),
            rotate_on_codes=rotate_codes,
            pause_451_s=pause451,
            failure_pause_threshold=fail_pause_n,
            attachments=self._get_attachments(),
            emails_per_minute=int(self.rate.get()),
            delay_min_s=dmin,
            delay_max_s=dmax,
            stop_event=self.stop_event,
            on_log=on_log,
            on_progress=on_progress,
            on_done=on_done,
            on_error=on_error,
        )
        self.sender.start()

        # persist
        self.cfg["active_account"] = self.active_account.get().strip()
        self.cfg["active_template"] = self.active_template.get().strip()
        try:
            self.cfg["rate_per_min"] = int(self.rate.get())
        except Exception:
            pass
        self.cfg["delay_min_s"] = dmin
        self.cfg["delay_max_s"] = dmax
        self.cfg["sender_name_override"] = int(self.sender_name_override.get())
        self.cfg["sender_name_global"] = self.sender_name_global.get()
        self.cfg["autopause_451"] = int(self.autopause_451.get())
        self.cfg["rotate_on_451"] = int(self.rotate_on_451.get())
        self.cfg["pause_451_s"] = pause451
        self.cfg["failure_pause_threshold"] = fail_pause_n
        self.cfg["auto_rotate_accounts"] = int(self.auto_rotate_accounts.get())
        self.cfg["rotate_account_on_codes"] = int(self.rotate_account_on_codes.get())
        self.cfg["rotate_account_codes"] = self.rotate_account_codes.get()
        self.cfg["use_proxies"] = int(self.use_proxies.get())
        self.cfg["rotate_proxies"] = int(self.rotate_proxies.get())
        try:
            self.cfg["rotate_proxy_every_n"] = int(self.rotate_proxy_every_n.get() or "0")
        except Exception:
            self.cfg["rotate_proxy_every_n"] = 0
        self.cfg["rotate_proxy_on_codes"] = int(self.rotate_proxy_on_codes.get())
        self.cfg["rotate_proxy_codes"] = self.rotate_proxy_codes.get()
        self.cfg["use_multiple_templates"] = int(self.use_multiple_templates.get())
        self.cfg["rotate_templates"] = int(self.rotate_templates.get())
        self.cfg["rotate_template_every_n"] = int(self.rotate_template_every_n.get() or "0")
        self.cfg["rotate_template_on_codes"] = int(self.rotate_template_on_codes.get())
        self.cfg["rotate_template_codes"] = self.rotate_template_codes.get()
        self.cfg["rotate_template_every_s"] = parse_float_relaxed(self.rotate_template_every_s.get(), 0.0)
        self.cfg["rotate_template_on_fail_streak"] = int(self.rotate_template_on_fail_streak.get())
        try:
            self.cfg["rotate_template_fail_streak_n"] = int(self.rotate_template_fail_streak_n.get() or "0")
        except Exception:
            self.cfg["rotate_template_fail_streak_n"] = 0
        try:
            self.cfg["rotate_every_n"] = int(self.rotate_every_n.get() or "0")
        except Exception:
            self.cfg["rotate_every_n"] = 0
        save_config(self.cfg)

    def continue_sending(self) -> None:
        if not self.csv_path or not self.emails:
            messagebox.showwarning("Проверка", "Сначала выбери CSV.")
            return
        if self.resume_index <= 0:
            messagebox.showinfo("Продолжить", "Нечего продолжать: позиция не сохранена.")
            return
        if self.sender is not None and self.sender.is_alive():
            return
        self._run_send(reset_log=False)

    def continue_with_other_account(self) -> None:
        if self.resume_index <= 0:
            messagebox.showinfo("Продолжить", "Нечего продолжать: позиция не сохранена.")
            return
        accounts = [a.get("name", "") for a in (self.cfg.get("accounts") or []) if a.get("name")]
        if not accounts:
            messagebox.showwarning("Аккаунты", "Список аккаунтов пуст.")
            return
        win = Toplevel(self.root)
        win.title("Продолжить с аккаунтом")
        win.geometry("420x140")
        v = StringVar(value=self.active_account.get() if self.active_account.get() in accounts else (accounts[0] if accounts else ""))
        frm = ttk.Frame(win)
        frm.pack(fill=BOTH, expand=True, padx=10, pady=10)
        r1 = ttk.Frame(frm)
        r1.pack(fill=BOTH, pady=6)
        ttk.Label(r1, text="Аккаунт:").pack(side=LEFT)
        cb = ttk.Combobox(r1, textvariable=v, values=accounts, state="readonly", width=28)
        cb.pack(side=LEFT, padx=6)

        def go():
            self.active_account.set(v.get())
            self.cfg["active_account"] = v.get()
            save_config(self.cfg)
            win.destroy()
            self._run_send(reset_log=False)

        r2 = ttk.Frame(frm)
        r2.pack(fill=BOTH, pady=6)
        ttk.Button(r2, text="Продолжить", command=go).pack(side=LEFT)
        ttk.Button(r2, text="Отмена", command=win.destroy).pack(side=LEFT, padx=6)

    def next_account(self) -> None:
        accounts = [a.get("name", "") for a in (self.cfg.get("accounts") or []) if isinstance(a, dict) and a.get("name")]
        if not accounts:
            messagebox.showwarning("Аккаунты", "Список аккаунтов пуст. Импортируй или добавь аккаунт.")
            return
        cur = self.active_account.get().strip()
        try:
            idx = accounts.index(cur)
            nxt = accounts[(idx + 1) % len(accounts)]
        except ValueError:
            nxt = accounts[0]

        self.active_account.set(nxt)
        self.cfg["active_account"] = nxt
        save_config(self.cfg)
        self._refresh_send_selectors()
        self._ui_log(now_ts(), "-", f"ACCOUNT: {nxt}")

    def export_log(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            initialfile=f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(["time", "email", "status"])
                for iid in self.tree.get_children():
                    vals = self.tree.item(iid, "values")
                    w.writerow(list(vals))
            messagebox.showinfo("Лог", f"Сохранено: {path}")
        except Exception as e:
            messagebox.showerror("Лог", f"Не удалось сохранить: {e}")

    def send_test_dialog(self) -> None:
        win = Toplevel(self.root)
        win.title("Тестовая отправка")
        win.geometry("520x140")
        v_email = StringVar(value="")
        frm = ttk.Frame(win)
        frm.pack(fill=BOTH, expand=True, padx=10, pady=10)
        r1 = ttk.Frame(frm)
        r1.pack(fill=BOTH, pady=6)
        ttk.Label(r1, text="Тестовый email:").pack(side=LEFT)
        ttk.Entry(r1, textvariable=v_email, width=40).pack(side=LEFT, padx=6)

        def go():
            e = normalize_email(v_email.get())
            if not e:
                messagebox.showwarning("Тест", "Укажи корректный email.")
                return
            win.destroy()
            self.send_test(e)

        r2 = ttk.Frame(frm)
        r2.pack(fill=BOTH, pady=6)
        ttk.Button(r2, text="Отправить", command=go).pack(side=LEFT)
        ttk.Button(r2, text="Отмена", command=win.destroy).pack(side=LEFT, padx=6)

    def send_test(self, email: str) -> None:
        # Uses current active account + template, sends exactly 1 email.
        if not self.active_account.get().strip() or not self.active_template.get().strip():
            messagebox.showwarning("Тест", "Выбери активный аккаунт и шаблон.")
            return
        acc_name = self.active_account.get().strip()
        tpl_name = self.active_template.get().strip()
        acc = next((a for a in (self.cfg.get("accounts") or []) if a.get("name") == acc_name), None)
        tpl = next((t for t in (self.cfg.get("templates") or []) if t.get("name") == tpl_name), None)
        if not acc or not tpl:
            messagebox.showerror("Тест", "Аккаунт или шаблон не найдены.")
            return

        smtp_cfg = SmtpConfig(
            host=str(acc.get("host", "")).strip(),
            port=int(acc.get("port", 587)),
            use_ssl=bool(int(acc.get("ssl", 0) or 0)),
            starttls=bool(int(acc.get("starttls", 1))),
            verify_tls=bool(int(acc.get("verify_tls", 1))),
            username=str(acc.get("username", "")).strip(),
            password=str(acc.get("password", "")),
            from_email=str(acc.get("from_email", "")).strip() or str(acc.get("username", "")).strip(),
            from_name=(self.sender_name_global.get().strip() if self.sender_name_override.get() else str(acc.get("sender_name", "")).strip()),
        )

        try:
            dmin = parse_float_relaxed(self.delay_min_s.get(), 0.0)
            dmax = parse_float_relaxed(self.delay_max_s.get(), 0.0)
            pause451 = parse_float_relaxed(self.pause_451_s.get(), 120.0)
        except Exception:
            dmin = dmax = 0.0
            pause451 = 120.0

        self._ui_log(now_ts(), email, "TEST: queued")

        stop_ev = threading.Event()

        def on_log(ts: str, em: str, status: str) -> None:
            self.root.after(0, lambda: self._ui_log(ts, em, "TEST: " + status if status == "SENT" else status))

        def on_progress(_sent: int, _failed: int, _rem: int, _done: int, _total: int) -> None:
            return

        def on_done(_paused: bool, _last_idx: int, _total: int) -> None:
            self.root.after(0, lambda: self._ui_log(now_ts(), email, "TEST: done"))

        def on_error(msg: str) -> None:
            self.root.after(0, lambda: (self._ui_log(now_ts(), email, f"TEST ERROR: {msg}"), messagebox.showerror("Тест", msg)))

        t = SenderThread(
            smtp_cfgs=[smtp_cfg],
            start_smtp_idx=0,
            rotate_every_n=0,
            proxies=[],
            use_proxies=False,
            rotate_proxy_every_n=0,
            rotate_proxy_on_codes=set(),
            templates=[
                TemplateConfig(
                    name=str(tpl.get("name", "")) or "template",
                    subject=str(tpl.get("subject", "")).strip(),
                    body=str(tpl.get("body", "")),
                    is_html=bool(int(tpl.get("is_html", 1) or 0)),
                    batch_n=int(tpl.get("batch_n", 0) or 0),
                )
            ],
            start_template_idx=0,
            rotate_template_every_n=0,
            rotate_template_on_codes=set(),
            rotate_template_every_s=0.0,
            rotate_template_fail_streak_n=0,
            emails=[email],
            autopause_451=bool(self.autopause_451.get()),
            rotate_on_451=False,
            rotate_on_codes=set(),
            pause_451_s=pause451,
            failure_pause_threshold=0,
            attachments=self._get_attachments(),
            emails_per_minute=1,
            delay_min_s=dmin,
            delay_max_s=dmax,
            stop_event=stop_ev,
            on_log=on_log,
            on_progress=on_progress,
            on_done=on_done,
            on_error=on_error,
        )
        t.start()

    def stop(self) -> None:
        self.stop_event.set()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    App().run()


if __name__ == "__main__":
    main()

