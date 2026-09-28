"""Saved email previews and explicit SMTP delivery with per-route state."""
from email.message import EmailMessage
from email import policy
from email.parser import BytesParser
from email.utils import formatdate
from email.headerregistry import Address
from email.errors import HeaderParseError
from pathlib import Path
import hashlib
import mimetypes
import os
import smtplib
import ssl
from .report_artifacts import sha256, verify_file, write_json

SENDER = 'bcc.notification.noreply@gmail.com'


def env_list(name, settings=None):
    values = os.environ if settings is None else settings
    return [x.strip() for x in values.get(name, '').split(',') if x.strip()]


def validate_recipients(recipients):
    """Validate plain envelope addresses without echoing private values."""
    if not recipients:
        raise ValueError('No configured recipients')
    for value in recipients:
        try:
            address = Address(addr_spec=value)
            if not address.username or not address.domain or any(c.isspace() for c in value):
                raise ValueError()
            # The sender does not enable SMTPUTF8.
            value.encode('ascii')
        except (ValueError, TypeError, UnicodeError, HeaderParseError) as exc:
            raise ValueError('Invalid recipient address; use plain ASCII mailbox@domain addresses') from exc


def attach_file(message, path):
    path = Path(path)
    main, sub = (mimetypes.guess_type(str(path))[0] or 'application/octet-stream').split('/', 1)
    message.add_attachment(path.read_bytes(), maintype=main, subtype=sub, filename=path.name)


def message(subject, body, attachments, recipients, *, cc=(), message_key=''):
    msg = EmailMessage()
    msg['From'] = SENDER
    msg['To'] = ', '.join(recipients)
    if cc:
        msg['Cc'] = ', '.join(cc)
    msg['Subject'] = subject
    msg['Date'] = formatdate(localtime=False)
    msg['Message-ID'] = '<' + hashlib.sha256(message_key.encode()).hexdigest() + '@bcc-reports.local>'
    msg.set_content(body)
    for path in attachments:
        attach_file(msg, path)
    return msg


def save_route(directory, manifest, channel, msg, envelope, *, max_bytes=25000000, test_recipient=''):
    routes = manifest.setdefault('routes', {})
    path = Path(directory) / f'{channel}.eml'
    if channel in routes:
        verify_file(path, routes[channel]['sha256'])
        return
    recipients = list(dict.fromkeys(x for x in envelope if x))
    if test_recipient:
        validate_recipients([test_recipient])
        recipients = [test_recipient]
        for header in ('To', 'Cc', 'Bcc'):
            if header in msg:
                del msg[header]
        msg['To'] = test_recipient
        msg.replace_header('Subject', '[PILOT] ' + str(msg['Subject']))
    raw = msg.as_bytes(policy=policy.SMTP)
    if len(raw) > max_bytes:
        raise RuntimeError(f'{channel}: serialized email exceeds configured size limit')
    path.write_bytes(raw)
    routes[channel] = {'state': 'prepared', 'file': path.name, 'sha256': sha256(path),
                       'envelope': recipients, 'bytes': len(raw), 'attempts': []}
    write_json(Path(directory) / 'manifest.json', manifest)


def deliver(directory, manifest, *, dry_run=True, resend_uncertain=False, smtp_factory=None,
            allowed_channels=None, settings=None):
    if dry_run:
        return
    routes = manifest.get('routes', {})
    pending = []
    for key, route in routes.items():
        if allowed_channels is not None and key not in allowed_channels:
            continue
        verify_file(Path(directory) / route['file'], route['sha256'])
        if route['state'] == 'sent':
            continue
        if route['state'] in ('submitting', 'uncertain') and not resend_uncertain:
            raise RuntimeError(f'{key}: SMTP outcome is uncertain; review before explicit resend')
        validate_recipients(route['envelope'])
        pending.append((key, route))
    if not pending:
        return
    password = (os.environ if settings is None else settings).get('SMTP_PASS')
    if not password:
        raise RuntimeError('SMTP_PASS not set')
    factory = smtp_factory or smtplib.SMTP
    with factory('smtp.gmail.com', 587, timeout=60) as smtp:
        smtp.ehlo()
        smtp.starttls(context=ssl.create_default_context())
        smtp.ehlo()
        smtp.login(SENDER, password)
        for key, route in pending:
            route['state'] = 'submitting'
            attempt = {'started': formatdate(localtime=False)}
            route['attempts'].append(attempt)
            write_json(Path(directory) / 'manifest.json', manifest)
            try:
                msg = BytesParser(policy=policy.default).parsebytes((Path(directory) / route['file']).read_bytes())
                refused = smtp.send_message(msg, from_addr=SENDER, to_addrs=route['envelope'])
                if refused:
                    raise RuntimeError(f'{key}: SMTP refused some recipients; accepted subset may have received mail')
            except Exception as exc:
                route['state'] = 'uncertain'
                attempt['outcome'] = type(exc).__name__
                write_json(Path(directory) / 'manifest.json', manifest)
                raise
            route['state'] = 'sent'
            attempt['outcome'] = 'accepted by SMTP'
            write_json(Path(directory) / 'manifest.json', manifest)
