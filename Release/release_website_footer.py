"""Version 1.0.0: append the standard website footer to a GitHub release body when it is missing.

Used by .github/workflows/release-website-footer.yml as a safety net for releases published
without the footer (including releases that bypass Publish-WorkplaceCloudHubRelease.ps1).
Reads the current body from a file and writes the new body only when a change is needed.
"""
import argparse
import json
import pathlib
import sys

SITE = 'https://workplacecloudhub.com/'
CATALOG = 'More tools: [workplacecloudhub.com/tools](https://workplacecloudhub.com/tools/)'
MAP_PATH = pathlib.Path(__file__).resolve().parent / 'website-footer-map.json'


def load_map(path=MAP_PATH):
    return json.loads(pathlib.Path(path).read_text(encoding='utf-8'))


def footer_line(tag, footer_map):
    """Return the footer line for a tag, or None when the tag must not link the website."""
    if any(tag.startswith(prefix) for prefix in footer_map['noProductLink']):
        return None
    matches = [p for p in footer_map['products'] if tag.startswith(p['prefix'])]
    if not matches:
        return CATALOG
    product = max(matches, key=lambda p: len(p['prefix']))
    return f"Product page: [{product['name']}](https://workplacecloudhub.com/tools/{product['slug']}/) | {CATALOG}"


def with_footer(tag, body, footer_map):
    """Return the body with the footer appended, or None when no change is needed."""
    body = (body or '').replace('\r\n', '\n')
    if SITE in body:
        return None
    line = footer_line(tag, footer_map)
    if line is None:
        return None
    head = body.rstrip()
    return (head + '\n\n' if head else '') + '---\n\n' + line + '\n'


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tag', required=True)
    parser.add_argument('--body-file', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args(argv)
    body = pathlib.Path(args.body_file).read_text(encoding='utf-8')
    updated = with_footer(args.tag, body, load_map())
    if updated is None:
        print(f'UNCHANGED {args.tag}')
        return 0
    pathlib.Path(args.output).write_text(updated, encoding='utf-8', newline='\n')
    print(f'UPDATED {args.tag}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
