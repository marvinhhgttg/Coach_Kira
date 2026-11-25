"""Coach Kira – Push an iPhone/iPad über Apple (APNs), ohne Zusatzpakete.
25.11.2026. Schlüssel bleibt lokal: ~/coach-kira-runner/apns/AuthKey_<KEYID>.p8
Konfiguration: ~/coach-kira-runner/apns/apns.json  {"team_id": "...", "topic": "de.reichpietsch.coachkira"}
Signieren mit openssl (ES256), Versand mit curl --http2.
"""
from __future__ import annotations

import base64
import json
import shutil
import subprocess
import time
from pathlib import Path

ORDNER = Path.home() / "coach-kira-runner" / "apns"
_jwt: tuple[float, str] | None = None


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _der_zu_rs(der: bytes) -> bytes:
    # ECDSA-Signatur (DER: 30 len 02 lr r 02 ls s) → r||s je 32 Byte
    i = 2 if der[1] < 0x80 else 2 + (der[1] & 0x7F)
    assert der[i] == 2
    lr = der[i + 1]; r = der[i + 2:i + 2 + lr]; i = i + 2 + lr
    assert der[i] == 2
    ls = der[i + 1]; s = der[i + 2:i + 2 + ls]
    return r.lstrip(b"\0").rjust(32, b"\0") + s.lstrip(b"\0").rjust(32, b"\0")


def konfig() -> dict | None:
    keys = sorted(ORDNER.glob("AuthKey_*.p8"))
    if not keys:
        return None
    k = {}
    try:
        k = json.loads((ORDNER / "apns.json").read_text())
    except Exception:
        pass
    datei = keys[-1]
    return {"key": datei, "key_id": k.get("key_id") or datei.stem.split("_", 1)[1],
            "team_id": k.get("team_id"), "topic": k.get("topic") or "de.reichpietsch.coachkira"}


def _openssl() -> str:
    return shutil.which("openssl") or "/usr/bin/openssl"


def jwt(k: dict) -> str:
    global _jwt
    if _jwt and time.time() - _jwt[0] < 40 * 60:
        return _jwt[1]
    kopf = _b64(json.dumps({"alg": "ES256", "kid": k["key_id"]}).encode())
    jetzt = int(time.time())
    inhalt = _b64(json.dumps({"iss": k["team_id"], "iat": jetzt}).encode())
    roh = f"{kopf}.{inhalt}".encode()
    r = subprocess.run([_openssl(), "dgst", "-sha256", "-sign", str(k["key"])], input=roh, capture_output=True, timeout=20)
    if r.returncode != 0:
        raise RuntimeError("openssl: " + r.stderr.decode()[:160])
    t = f"{kopf}.{inhalt}.{_b64(_der_zu_rs(r.stdout))}"
    _jwt = (time.time(), t)
    return t


def senden(ziele: list[dict], titel: str, text: str, kennung: str = "") -> list[str]:
    """Schickt eine Mitteilung an alle Ziele. Gibt Kennungen zurück, die Apple als ungültig meldet."""
    k = konfig()
    if not k or not k.get("team_id") or not ziele:
        return []
    weg = []
    nutz = json.dumps({"aps": {"alert": {"title": titel[:120], "body": text[:900]}, "sound": "default",
                               "thread-id": "coach-kira"}}, ensure_ascii=False)
    for z in ziele:
        tok = str(z.get("token") or "")
        if not tok:
            continue
        host = "api.sandbox.push.apple.com" if z.get("umgebung") == "development" else "api.push.apple.com"
        kopf = ["-H", f"authorization: bearer {jwt(k)}", "-H", f"apns-topic: {k['topic']}",
                "-H", "apns-push-type: alert", "-H", "apns-priority: 10"]
        if kennung:
            kopf += ["-H", f"apns-collapse-id: {kennung[:60]}"]
        r = subprocess.run(["/usr/bin/curl", "-sS", "--http2", "-o", "-", "-w", "\n%{http_code}", *kopf,
                            "--data-binary", nutz, f"https://{host}/3/device/{tok}"],
                           capture_output=True, text=True, timeout=30)
        code = (r.stdout.strip().rsplit("\n", 1) + [""])[-1]
        if code in ("400", "410") and ("BadDeviceToken" in r.stdout or "Unregistered" in r.stdout or "DeviceTokenNotForTopic" in r.stdout):
            weg.append(tok)
        elif code != "200":
            raise RuntimeError(f"APNs {code}: {r.stdout.strip()[:160]} {r.stderr.strip()[:80]}")
    return weg
