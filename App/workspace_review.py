"""Content-based optimistic concurrency and confirmation tied to a review."""
import hashlib
import hmac
import json
from pathlib import Path
import secrets

from workflow_support import load_settings

_SECRET = secrets.token_bytes(32)


def workspace_revision(root, file_overrides=None, *, force=False):
    root = Path(root).resolve()
    settings = load_settings(root)
    paths = [root / 'settings.local.json', root / 'Data' / 'Custom Certificates.json',
             root / 'Data' / 'Certificate Inventory.json', root / 'Data' / 'History Reset.json',
             root / 'Data' / 'Certificate Folder Moves.json']
    paths.extend(Path(value).resolve() for value in settings.get('workbook_paths', []) if value)
    for folder in (root / 'Certificates', root / 'Data' / 'Attachments', root / 'App' / 'templates' / 'certificate-types'):
        paths.extend(path for path in sorted(folder.rglob('*')) if path.is_file() and not path.name.startswith('.'))
    result = hashlib.sha256()
    for path in dict.fromkeys(paths):
        result.update(str(path).encode('utf-8'))
        if file_overrides and str(path.resolve()) in file_overrides:
            result.update(hashlib.sha256(file_overrides[str(path.resolve())]).digest())
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


def require_revision(root, expected, file_overrides=None):
    if not isinstance(expected, str) or not expected or not hmac.compare_digest(workspace_revision(root, file_overrides, force=True), expected):
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
    check = dict(payload, confirmed=False)
    result = action(check)
    # A no-warning operation has already completed. Actions are deliberately
    # called once in that case, not duplicated after a successful issuance.
    if not result.get('requires_confirmation'):
        return result
    revision = workspace_revision(root)
    token = review_token(endpoint, payload, revision, result.get('warnings', []))
    if payload.get('confirmed') and isinstance(payload.get('review_token'), str) and hmac.compare_digest(payload['review_token'], token):
        reason = str(payload.get('override_reason') or '').strip()
        if len(reason) > 2000:
            raise ValueError('The warning confirmation note must be 2,000 characters or fewer.')
        # The reviewed token proves which details and warnings were accepted.
        # Retain notes from older clients, without requiring the user to type one.
        return action(dict(payload, expected_revision=revision,
                           override_reason=reason or 'Warnings reviewed; user chose Continue anyway.'))
    result.update(review_token=token, workspace_revision=revision)
    return result
