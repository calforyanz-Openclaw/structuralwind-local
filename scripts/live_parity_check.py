#!/usr/bin/env python3
import hashlib
import json
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path

LIVE_BASE = 'https://structuralwind.com'
LOCAL_BASE = 'http://127.0.0.1:8018'
PAGES = [
    ('index', '/', '/'),
    ('about', '/about.html', '/public/about.html'),
    ('script', '/script.js', '/public/script.js'),
]
CLOUD_STRINGS = [
    'firebase',
    'sign in with google',
    'sign in with microsoft',
    'create-checkout-session',
    'create-portal-session',
    'shared-projects',
    'api-keys',
    'enterprise-enquiry',
]
API_ROUTE_RE = re.compile(r"/api/[A-Za-z0-9\-?=]+")

class ScriptParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.scripts = []
    def handle_starttag(self, tag, attrs):
        if tag.lower() != 'script':
            return
        for k, v in attrs:
            if k.lower() == 'src' and v:
                self.scripts.append(v)

class TitleParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_title = False
        self.title = ''
    def handle_starttag(self, tag, attrs):
        if tag.lower() == 'title':
            self.in_title = True
    def handle_endtag(self, tag):
        if tag.lower() == 'title':
            self.in_title = False
    def handle_data(self, data):
        if self.in_title:
            self.title += data

def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'OpenClaw parity check'})
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.read().decode('utf-8', 'ignore')

def sha(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:16]

def title_of(html):
    p = TitleParser()
    p.feed(html)
    return p.title.strip()

def scripts_of(html):
    p = ScriptParser()
    p.feed(html)
    return p.scripts

def route_set(text):
    return sorted(set(API_ROUTE_RE.findall(text)))

def count_cloud_strings(text):
    lower = text.lower()
    return {s: lower.count(s) for s in CLOUD_STRINGS}

def main():
    out = []
    out.append(f'# Live Parity Report\n')
    out.append(f'- Generated: {datetime.now(timezone.utc).isoformat()}')
    out.append(f'- Live base: `{LIVE_BASE}`')
    out.append(f'- Local base: `{LOCAL_BASE}`\n')

    fetched = {}
    for label, live_path, local_path in PAGES:
        live_url = urllib.parse.urljoin(LIVE_BASE, live_path)
        local_url = urllib.parse.urljoin(LOCAL_BASE, local_path)
        live = fetch(live_url)
        local = fetch(local_url)
        fetched[label] = {'live': live, 'local': local, 'live_url': live_url, 'local_url': local_url}

        out.append(f'## {label}')
        out.append(f'- live url: `{live_url}`')
        out.append(f'- local url: `{local_url}`')
        out.append(f'- live sha16: `{sha(live)}`')
        out.append(f'- local sha16: `{sha(local)}`')
        out.append(f'- identical: `{live == local}`')
        if label != 'script':
            out.append(f'- live title: `{title_of(live)}`')
            out.append(f'- local title: `{title_of(local)}`')
            live_scripts = scripts_of(live)
            local_scripts = scripts_of(local)
            out.append(f'- live script refs: `{len(live_scripts)}`')
            out.append(f'- local script refs: `{len(local_scripts)}`')
            missing_from_local = [s for s in live_scripts if s not in local_scripts]
            added_in_local = [s for s in local_scripts if s not in live_scripts]
            out.append(f'- missing script refs in local: `{len(missing_from_local)}`')
            out.append(f'- added script refs in local: `{len(added_in_local)}`')
        else:
            live_routes = route_set(live)
            local_routes = route_set(local)
            only_live = [r for r in live_routes if r not in local_routes]
            only_local = [r for r in local_routes if r not in live_routes]
            out.append(f'- live api routes found: `{len(live_routes)}`')
            out.append(f'- local api routes found: `{len(local_routes)}`')
            out.append(f'- routes only in live script: `{only_live}`')
            out.append(f'- routes only in local script: `{only_local}`')
            out.append(f'- live cloud string counts: `{json.dumps(count_cloud_strings(live), ensure_ascii=False)}`')
            out.append(f'- local cloud string counts: `{json.dumps(count_cloud_strings(local), ensure_ascii=False)}`')
        out.append('')

    index_live = fetched['index']['live'].lower()
    index_local = fetched['index']['local'].lower()
    about_live = fetched['about']['live'].lower()
    about_local = fetched['about']['local'].lower()
    script_live = fetched['script']['live'].lower()
    script_local = fetched['script']['local'].lower()

    findings = []
    if 'firebase-app-compat.js' in index_live and 'firebase-app-compat.js' not in index_local:
        findings.append('Local index has removed Firebase SDK includes while live still serves them.')
    if 'sign in with google' in index_live and 'sign in with google' not in index_local:
        findings.append('Local auth overlay is localized away from cloud sign-in.')
    if 'cloud authentication and billing have been removed from this workspace.' in about_local:
        findings.append('About page now explicitly documents local-only behavior.')
    if 'create-checkout-session' in script_live and 'create-checkout-session' in script_local:
        findings.append('Checkout route references still exist in local script and should be cleaned further if full local-only parity is desired.')
    if 'shared-projects' in script_local:
        findings.append('Shared-project route references still exist in local script; behavior is stubbed but code remains.')
    if 'api-keys' in script_local:
        findings.append('API key overlay code still exists in local script; behavior is stubbed but not removed.')

    out.append('## Findings')
    for item in findings:
        out.append(f'- {item}')
    out.append('')
    out.append('## Next recommended checks')
    out.append('- Remove dead cloud-only overlay/rendering code paths now that local stubs are in place.')
    out.append('- Use browser screenshots to compare the live and local landing pages visually.')
    out.append('- Build one or two shared numeric test cases and compare the calculated outputs between live and local.')

    report = '\n'.join(out) + '\n'
    report_path = Path('/Users/jasonjia-claw/.openclaw/workspace/structuralwind-local/notes/live-parity-report.md')
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(report, encoding='utf-8')
    print(str(report_path))

if __name__ == '__main__':
    sys.exit(main() or 0)
