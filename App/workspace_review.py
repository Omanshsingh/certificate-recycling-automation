"""Content-based optimistic concurrency and confirmation tied to a review."""
import hashlib
import hmac
import json
from pathlib import Path
import re
import secrets

from workflow_support import load_settings

_SECRET = secrets.token_bytes(32)


def workspace_revision(root, file_overrides=None, *, force=False):
    root = Path(root).resolve()
    settings = load_settings(root)
    paths = [root / 'settings.local.json', root / 'Data' / 'Custom Certificates.json',
             root / 'Data' / 'Certificate Inventory.json', root / 'Data' / 'History Reset.json',
             root / 'Data' / 'Certificate Folder Moves.json',
             root / 'Data' / 'Issued Excel Change Reviews.json']
    paths.extend(Path(value).resolve() for value in settings.get('workbook_paths', []) if value)
    for folder in (root / 'Certificates', root / 'Data' / 'Attachments', root / 'App' / 'templates' / 'certificate-types'):
        paths.extend(path for path in sorted(folder.rglob('*')) if path.is_file() and not path.name.startswith('.'))
    result = hashlib.sha256()
    for path in dict.fromkeys(paths):
        name = str(path)
        result.update(name.encode('utf-8'))
        if file_overrides and name in file_overrides:
            result.update(hashlib.sha256(file_overrides[name]).digest())
            continue
        try:
            from file_fingerprints import fingerprint
            digest, _ = fingerprint(path, force=force)
            result.update(bytes.fromhex(digest))
        except FileNotFoundError:
            result.update(b'missing')
        except OSError as exc:
            raise ValueError(f'Close Excel and retry; {path.name} could not be checked for changes.') from exc
    return result.hexdigest()


def require_revision(root, expected, file_overrides=None, *, force=True):
    if (not isinstance(expected, str) or not re.fullmatch(r'[a-f0-9]{64}', expected)
            or not hmac.compare_digest(workspace_revision(root, file_overrides, force=force), expected)):
        raise ValueError('The files changed since you opened these details. Refresh Excel and review the current details before saving. Your entered details have not been saved.')


def review_token(endpoint, payload, revision, warnings):
    clean = {key: value for key, value in payload.items()
             if key not in {'confirmed', 'review_token', 'override_reason', 'expected_revision'}}
    raw = json.dumps([endpoint, clean, revision, warnings], sort_keys=True, default=str, ensure_ascii=False).encode('utf-8')
    return hmac.new(_SECRET, raw, hashlib.sha256).hexdigest()


def reviewed_operation(root, endpoint, payload, action):
    """HTTP boundary: never accept a bare confirmed=True as prior review."""
    expected = payload.get('expected_revision')
    require_revision(root, expected)
    check = dict(payload, confirmed=False, review_only=True)
    result = action(check)
    if not result.get('requires_confirmation'):
        return result
    # A file can change while the warnings are being computed. Keep the
    # review tied to the original revision rather than blessing newer data.
    require_revision(root, expected, force=False)
    revision = expected
    token = review_token(endpoint, payload, revision, result.get('warnings', []))
    submitted_token = payload.get('review_token')
    if (payload.get('confirmed') is True and isinstance(submitted_token, str)
            and submitted_token.isascii() and hmac.compare_digest(submitted_token, token)):
        reason = str(payload.get('override_reason') or '').strip()
        if len(reason) > 2000:
            raise ValueError('The warning confirmation note must be 2,000 characters or fewer.')
        # The reviewed token proves which details and warnings were accepted.
        # Retain notes from older clients, without requiring the user to type one.
        default_reason = 'Warnings reviewed; user chose Continue anyway.' if result.get('warnings') else 'Creation confirmed.'
        return action(dict(payload, expected_revision=revision, review_only=False,
                           override_reason=reason or default_reason))
    result.update(review_token=token, workspace_revision=revision)
    return result
