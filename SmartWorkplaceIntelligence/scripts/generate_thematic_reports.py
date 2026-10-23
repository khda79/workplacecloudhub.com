"""Generate report-only PBIR derivatives, v1.2.3; never edit or deploy the master.

The default operation is a read-only plan. --write creates a NEW output folder.
Stdout is a machine-readable JSON receipt. No network or Desktop calls occur.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import uuid


VERSION = "1.2.3"
DEFAULT_MANIFEST = Path(__file__).resolve().parents[1] / "config/thematic-reports.json.txt"
PAGE_REFERENCES = {"navigationSection", "tooltipSection", "drillthroughSection", "pageName"}


class GenerationError(ValueError):
    """Invalid source, unsafe destination, or unresolved dependency."""


def require(condition, message):
    if not condition:
        raise GenerationError(message)


def reject_duplicates(pairs):
    result = {}
    for key, value in pairs:
        require(key not in result, f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def decode(data):
    return json.loads(data.decode("utf-8-sig"), object_pairs_hook=reject_duplicates)


def encode(value):
    return (json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def digest(data):
    return hashlib.sha256(data).hexdigest()


def linked(path):
    if path.is_symlink() or (hasattr(path, "is_junction") and path.is_junction()):
        return True
    # Python 3.10/3.11 lack Path.is_junction on Windows.
    try:
        return getattr(path.lstat(), 'st_reparse_tag', 0) in {0xA0000003, 0xA000000C}
    except FileNotFoundError:
        return False


def checked_path(path):
    path = Path(os.path.abspath(path))
    for part in (path, *path.parents):
        require(not linked(part), f"Linked paths are not supported: {part}")
    return path


def under(path, parent):
    return path == parent or parent in path.parents


def snapshot(root):
    """Capture report definitions/resources only; never read local caches."""
    root = checked_path(root)
    require(root.is_dir() and root.name.endswith(".Report"), "Source must be a .Report directory")
    payload = {}
    def read_error(error):
        raise error
    for folder, directories, files in os.walk(root, followlinks=False, onerror=read_error):
        directories[:] = sorted(d for d in directories if d not in {".pbi", ".git"})
        for directory in directories:
            require(not linked(Path(folder) / directory), "Linked report directories are not supported")
        for name in sorted(files):
            path = Path(folder) / name
            require(not linked(path), "Linked report files are not supported")
            rel = path.relative_to(root).as_posix()
            require(rel.startswith(("definition/", "StaticResources/", "CustomVisuals/"))
                    or rel in {"definition.pbir", ".platform"}, f"Unsupported report file: {rel}")
            payload[rel] = path.read_bytes()
    for required in ("definition.pbir", ".platform", "definition/report.json",
                     "definition/version.json", "definition/pages/pages.json"):
        require(required in payload, f"Missing source file: {required}")
    return payload


def fingerprints(payload):
    return {key: digest(data) for key, data in sorted(payload.items())}


def literal(value):
    if isinstance(value, str):
        return value
    require(isinstance(value, dict), "Unsupported dependency expression")
    expr = value.get("expr", value)
    require(set(expr) == {"Literal"}, "Dynamic dependency expressions need an explicit adapter")
    token = expr["Literal"]["Value"]
    require(isinstance(token, str) and token.startswith("'") and token.endswith("'"),
            "Dependency must be a text literal")
    return token[1:-1].replace("''", "'")


def walk(value):
    if isinstance(value, dict):
        for key, child in value.items():
            yield key, child
            yield from walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from walk(child)


def references(document):
    pages, bookmarks, groups = set(), set(), set()
    for key, value in walk(document):
        if key in PAGE_REFERENCES:
            pages.add(literal(value))
        elif key == "bookmark":
            bookmarks.add(literal(value))
        elif key == "bookmarkGroup":
            groups.add(literal(value))
    return pages - {""}, bookmarks - {""}, groups - {""}


def report_file_name(spec):
    name = spec.get("fileName", f"SWI-{spec['key']}")
    require(isinstance(name, str) and bool(name) and name == name.strip()
            and not re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
            and not name.endswith('.') and len(name) <= 120
            and name.split('.')[0].upper() not in
            {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1, 10)),
             *(f'LPT{i}' for i in range(1, 10))}, "Invalid report file name")
    return name


def read_manifest(path, all_pages):
    data = decode(path if isinstance(path, bytes) else Path(path).read_bytes())
    require(data.get("version") == 1, "Unsupported thematic manifest version")
    require(data.get('outputLayout', 'per-theme') in {'per-theme', 'root-projects'},
            "Unsupported output layout")
    reports = data.get("reports", [])
    require(bool(reports), "No thematic reports configured")
    keys, orders, assigned, filenames = set(), set(), set(), set()
    for spec in reports:
        key, order, pages = spec["key"], spec["order"], spec["pages"]
        require(isinstance(key, str) and re.fullmatch(r"[a-z][a-z0-9-]*", key), "Invalid theme key")
        require(key not in keys and order not in orders, "Duplicate theme key or order")
        require(isinstance(order, int) and not isinstance(order, bool) and order > 0, "Invalid theme order")
        require(isinstance(spec.get("displayName"), str) and bool(spec["displayName"].strip()),
                "Missing report display name")
        filename = report_file_name(spec).casefold()
        require(filename not in filenames, "Duplicate report file name")
        filenames.add(filename)
        require(bool(pages) and len(pages) == len(set(pages)), f"Empty or duplicate pages: {key}")
        require(set(pages) <= all_pages, f"Unknown configured pages: {sorted(set(pages) - all_pages)}")
        require(not assigned.intersection(pages), "A principal page is assigned to multiple themes")
        support = spec.get("supportingPages", [])
        require(isinstance(support, list) and len(support) == len(set(support)), "Duplicate supporting pages")
        require(set(support) <= all_pages and not set(support).intersection(pages),
                "Supporting pages must exist and differ from principal pages")
        keys.add(key)
        orders.add(order)
        assigned.update(pages)
    require(assigned == all_pages, f"Unclassified source pages: {sorted(all_pages - assigned)}")
    repairs = data.get("orphanBookmarkTargets", {})
    require(isinstance(repairs, dict), "Invalid orphan bookmark target allowlist")
    for name, targets in repairs.items():
        require(isinstance(name, str) and isinstance(targets, list) and targets
                and all(isinstance(t, str) for t in targets) and len(targets) == len(set(targets)),
                "Invalid orphan bookmark target allowlist entry")
    return data


def bookmark_metadata(document, retained):
    def prune(item):
        if isinstance(item, str):
            return item if item in retained else None
        require(isinstance(item, dict) and "children" in item, "Unsupported bookmark metadata item")
        children = [filtered for child in item["children"] if (filtered := prune(child)) is not None]
        if not children:
            return None
        result = copy.deepcopy(item)
        result["children"] = children
        return result
    result = copy.deepcopy(document)
    result["items"] = [filtered for item in document.get("items", [])
                       if (filtered := prune(item)) is not None]
    return result


def metadata_members(document):
    members, groups = set(), {}
    def collect(item):
        if isinstance(item, str):
            require(item not in members, f"Duplicate bookmark metadata member: {item}")
            members.add(item)
            return {item}
        require(isinstance(item, dict) and "children" in item, "Unsupported bookmark metadata")
        require(item["name"] not in groups, "Duplicate bookmark group")
        result = set()
        for child in item["children"]:
            result.update(collect(child))
        groups[item["name"]] = result
        return result
    for item in document.get("items", []):
        collect(item)
    return members, groups


def simplify_navigation(payload, selected, policy):
    """Adapt derivative bytes only; preserve content views and protected buttons."""
    require(policy.get('mode') == 'tabs', 'Unsupported navigation policy')
    protected = set(policy.get('protectedButtonPages', []))
    require(all(isinstance(p, str) for p in protected), 'Invalid protected button pages')
    output = dict(payload)
    removed = {page: set() for page in selected}
    changes = {'mode': 'tabs', 'removedNavigationVisuals': [], 'removedDrillthroughBindings': [],
               'addedExplorerSelectors': [], 'reflowedSlicers': [], 'prunedBookmarks': [],
               'layoutAdjustments': []}
    bookmarks = {decode(data)['name']: decode(data) for path, data in payload.items()
                 if path.endswith('.bookmark.json')}
    dead_filters = set()
    for page in sorted(selected):
        path = f'definition/pages/{page}/page.json'
        document = decode(output[path])
        if document.get('pageBinding', {}).get('type') == 'Drillthrough':
            dead_filters.update(p['boundFilter'] for p in document['pageBinding'].get('parameters', [])
                                if p.get('boundFilter'))
            document.pop('pageBinding')
            changes['removedDrillthroughBindings'].append(page)
        # Principal pages are normal visible tabs, including former detail pages.
        document.pop('visibility', None)
        output[path] = encode(document)
        for visual_path in list(output):
            if not visual_path.startswith(f'definition/pages/{page}/visuals/') or not visual_path.endswith('/visual.json'):
                continue
            visual = decode(output[visual_path])
            links = visual.get('visual', {}).get('visualContainerObjects', {}).get('visualLink', [])
            remove_action = False
            for link in links:
                props = link.get('properties', {})
                action = literal(props['type']) if 'type' in props else ''
                remove_action |= action in {'Back', 'PageNavigation', 'Drillthrough'}
                if action == 'Bookmark':
                    target = literal(props['bookmark'])
                    require(target in bookmarks, f'Missing bookmark: {target}')
                    destination = bookmarks[target].get('explorationState', {}).get('activeSection')
                    remove_action |= bool(destination and destination != page)
            if not remove_action or page in protected:
                continue
            if visual.get('visual', {}).get('visualType') == 'actionButton':
                removed[page].add(visual['name'])
                del output[visual_path]
                changes['removedNavigationVisuals'].append({'page': page, 'visual': visual['name']})
            else:
                visual['visual']['visualContainerObjects'].pop('visualLink')
                output[visual_path] = encode(visual)

    # Only references to explicitly removed navigation/filter IDs are pruned.
    # Unknown references still fail the normal dependency checks below.
    def prune_filters(value):
        if isinstance(value, dict):
            for key, child in value.items():
                if key == 'filters' and isinstance(child, list):
                    value[key] = [f for f in child if not isinstance(f, dict) or f.get('name') not in dead_filters]
                prune_filters(value[key])
        elif isinstance(value, list):
            for child in value:
                prune_filters(child)

    removed_bookmarks = set()
    for path in list(output):
        if not path.endswith('.json') or not path.startswith('definition/'):
            continue
        document = decode(output[path])
        before = copy.deepcopy(document)
        if path.endswith('.bookmark.json'):
            state = document.get('explorationState', {})
            sections = state.get('sections', {})
            for page, section in sections.items():
                for field in ('visualContainers', 'visualContainerGroups'):
                    for name in removed.get(page, set()):
                        section.get(field, {}).pop(name, None)
            scope_removed = set().union(*(removed.get(page, set()) for page in sections)) if sections else set()
            options = document.get('options', {})
            if 'targetVisualNames' in options:
                old_targets = options['targetVisualNames']
                options['targetVisualNames'] = [v for v in old_targets if v not in scope_removed]
                if old_targets and not options['targetVisualNames']:
                    require(set(old_targets) <= scope_removed, 'Unsafe empty bookmark scope')
                    removed_bookmarks.add(document['name'])
                    del output[path]
                    changes['prunedBookmarks'].append(document['name'])
                    continue
        parts = path.split('/')
        if path.startswith('definition/pages/') and len(parts) > 3 and parts[2] in selected:
            page_removed = removed[parts[2]]
            if 'visualInteractions' in document:
                document['visualInteractions'] = [i for i in document['visualInteractions']
                    if i['source'] not in page_removed and i['target'] not in page_removed]
        prune_filters(document)
        if document != before:
            output[path] = encode(document)
    metadata = 'definition/bookmarks/bookmarks.json'
    if removed_bookmarks and metadata in output:
        remaining = {decode(b)['name'] for p, b in output.items() if p.endswith('.bookmark.json')}
        output[metadata] = encode(bookmark_metadata(decode(output[metadata]), remaining))

    for spec in policy.get('explorerSelectors', []):
        page = spec['page']
        if page not in selected:
            continue
        prefix = f'definition/pages/{page}/visuals/'
        paths = [p for p in output if p.startswith(prefix) and p.endswith('/visual.json')]
        field = {'Column': {'Expression': {'SourceRef': {'Entity': spec['entity']}}, 'Property': spec['column']}}
        for p in paths:
            visual = decode(output[p]).get('visual', {})
            if visual.get('visualType') == 'slicer' and any(v == field for k, v in walk(visual.get('query', {})) if k == 'field'):
                break
        else:
            template_path = prefix + spec['templateVisual'] + '/visual.json'
            require(template_path in output, 'Missing Explorer selector template')
            selector = decode(output[template_path])
            require(selector['visual']['visualType'] == 'slicer', 'Explorer template must be a slicer')
            selector['name'] = uuid.uuid5(uuid.NAMESPACE_URL, f"swi-explorer:{page}:{spec['entity']}:{spec['column']}").hex[:20]
            selector['position'] = dict(spec['position'], z=27000, tabOrder=27000)
            selector.pop('parentGroupName', None)
            selector.pop('filterConfig', None)
            visual = selector['visual']
            visual.pop('syncGroup', None)
            visual['query'] = {'queryState': {'Values': {'projections': [{'field': field,
                'queryRef': spec['entity'] + '.' + spec['column'], 'nativeQueryRef': spec['column'], 'active': True}]}}}
            for properties in visual.get('objects', {}).get('general', []):
                properties.get('properties', {}).pop('filter', None)
            for header in visual.get('objects', {}).get('header', []):
                header['properties']['text'] = {'expr': {'Literal': {'Value': "'" + spec['column'].replace("'", "''") + "'"}}}
            for p in paths:
                candidate = decode(output[p])
                reflow = spec.get('reflowSlicers', {}).get(candidate['name'])
                if reflow:
                    require(candidate['visual']['visualType'] == 'slicer', 'Reflow target must be a slicer')
                    candidate['position'].update(reflow)
                    output[p] = encode(candidate)
                    changes['reflowedSlicers'].append({'page': page, 'visual': candidate['name']})
            selector_path = prefix + selector['name'] + '/visual.json'
            require(selector_path not in output, 'Duplicate Explorer selector')
            output[selector_path] = encode(selector)
            changes['addedExplorerSelectors'].append({'page': page, 'visual': selector['name'],
                                                     'entity': spec['entity'], 'column': spec['column']})
    for patch in policy.get('layoutAdjustments', []):
        page = patch['page']
        if page not in selected:
            continue
        path = f"definition/pages/{page}/visuals/{patch['visual']}/visual.json"
        require(path in output, 'Missing layout adjustment visual')
        document = decode(output[path])
        require(document.get('visual', {}).get('visualType') == 'slicer',
                'Layout adjustment is restricted to slicers')
        before = dict(document['position'])
        require(all(before.get(k) == v for k, v in patch['expectedPosition'].items()),
                'Layout source geometry changed; review adjustment')
        document['position'].update(patch['position'])
        position = document['position']
        page_document = decode(output[f'definition/pages/{page}/page.json'])
        require(position['x'] >= 0 and position['y'] >= 0
                and position['x'] + position['width'] <= page_document['width']
                and position['y'] + position['height'] <= page_document['height'],
                'Adjusted slicer exceeds the page')
        for other_path, data in output.items():
            if other_path == path or not other_path.startswith(f'definition/pages/{page}/visuals/') or not other_path.endswith('/visual.json'):
                continue
            other = decode(data)
            if not other.get('visual') or other.get('parentGroupName') != document.get('parentGroupName') or 'position' not in other:
                continue
            box = other['position']
            overlap = (position['x'] < box['x'] + box['width'] and position['x'] + position['width'] > box['x']
                       and position['y'] < box['y'] + box['height'] and position['y'] + position['height'] > box['y'])
            require(not overlap, f"Adjusted slicer overlaps a sibling visual: {other['name']}")
        output[path] = encode(document)
        changes['layoutAdjustments'].append({'page': page, 'visual': document['name'],
                                             'before': before, 'after': dict(position)})
    # A protected button may not silently change, even through a generic adapter.
    for path, data in payload.items():
        parts = path.split('/')
        if len(parts) > 4 and parts[2] in protected and parts[2] in selected and path.endswith('/visual.json'):
            if decode(data).get('visual', {}).get('visualType') == 'actionButton':
                require(output.get(path) == data, 'Protected button changed')
    return output, changes


def build(source, payload, manifest, theme, report_dir, semantic_model_id=None):
    """Pure generation from captured bytes; report_dir is the final binding location."""
    original_payload = payload
    page_meta = decode(payload["definition/pages/pages.json"])
    order = page_meta["pageOrder"]
    require(len(order) == len(set(order)) and bool(order), "Duplicate or empty master page order")
    physical = {p.split('/')[2] for p in payload if p.startswith("definition/pages/") and p.endswith('/page.json')}
    require(set(order) == physical, "Master page order does not match page directories")
    manifest_data = read_manifest(manifest, set(order))
    specs = [s for s in manifest_data["reports"] if s["key"] == theme]
    require(len(specs) == 1, f"Unknown report theme: {theme}")
    spec = specs[0]
    support = spec.get("supportingPages", [])
    selected = set(spec["pages"]) | set(support)
    navigation = manifest_data.get('navigation')
    navigation_changes = None
    if navigation:
        require(not support, 'Tab navigation must not retain hidden supporting pages')
        payload, navigation_changes = simplify_navigation(payload, selected, navigation)
    page_payload = {p: b for p, b in payload.items()
                    if p.startswith("definition/pages/") and len(p.split('/')) > 3
                    and p.split('/')[2] in selected}
    for page in selected:
        document = decode(page_payload[f"definition/pages/{page}/page.json"])
        require(document["name"] == page, "Page directory/name mismatch")
    page_refs, bookmark_refs, group_refs = set(), set(), set()
    for path, data in page_payload.items():
        if path.endswith('.json'):
            p, b, g = references(decode(data))
            page_refs.update(p)
            bookmark_refs.update(b)
            group_refs.update(g)
    require(page_refs <= selected,
            f"External page dependencies need an explicit adapter: {sorted(page_refs - selected)}")
    documents = {decode(b)["name"]: (p, decode(b)) for p, b in payload.items()
                 if p.startswith("definition/bookmarks/") and p.endswith('.bookmark.json')}
    require(len(documents) == sum(p.endswith('.bookmark.json') for p in payload), "Duplicate bookmark names")
    retained = set()
    for name, (path, document) in documents.items():
        state = document.get("explorationState", {})
        sections = set(state.get("sections", {}))
        active = state.get("activeSection")
        if active in selected or sections.intersection(selected):
            require(sections <= selected and (not active or active in selected),
                    f"Cross-theme bookmark requires explicit handling: {name}")
            retained.add(name)
    require(bookmark_refs <= retained, f"Missing or external bookmarks: {sorted(bookmark_refs - retained)}")
    source_bm_meta = decode(payload.get("definition/bookmarks/bookmarks.json", b'{"items":[]}'))
    meta = bookmark_metadata(source_bm_meta, retained)
    members, groups = metadata_members(meta)
    require(members == retained, "Bookmark metadata and retained bookmark files differ")
    require(group_refs <= groups.keys(), "Missing or external bookmark navigator group")
    # A navigator referring to a partially retained group would silently change its meaning.
    _, original_groups = metadata_members(source_bm_meta)
    for group in group_refs:
        require(original_groups[group] == groups[group], "Bookmark group spans multiple themes")
    visual_ids = {page: set() for page in selected}
    for path, data in page_payload.items():
        parts = path.split('/')
        if path.endswith('/visual.json'):
            visual = decode(data)
            require(visual['name'] == parts[-2], "Visual directory/name mismatch")
            require(visual['name'] not in visual_ids[parts[2]], "Duplicate visual ID")
            visual_ids[parts[2]].add(visual['name'])
        elif path.endswith('/visualGroup.json'):
            visual_ids[parts[2]].add(decode(data)['name'])
    for path, data in page_payload.items():
        if not path.endswith('.json'):
            continue
        document = decode(data)
        page = path.split('/')[2]
        if document.get('parentGroupName'):
            require(document['parentGroupName'] in visual_ids[page], "Missing parent visual group")
        for interaction in document.get('visualInteractions', []):
            require({interaction['source'], interaction['target']} <= visual_ids[page],
                    "Interaction refers to a missing visual")
    retained_payload = {}
    repairs = []
    for name in retained:
        path, document = documents[name]
        p, b, g = references(document)
        require(p <= selected and b <= retained and g <= groups.keys(), "Unresolved bookmark dependency")
        state = document.get("explorationState", {})
        scope = set(state.get("sections", {}))
        targets = set(document.get("options", {}).get("targetVisualNames", []))
        known = set().union(*(visual_ids[page] for page in scope)) if scope else set()
        missing = targets - known
        allowed = set(manifest_data.get("orphanBookmarkTargets", {}).get(name, []))
        require(missing <= allowed, f"Bookmark targets missing visuals: {name}: {sorted(missing - allowed)}")
        if missing:
            # Only reviewed dead target IDs are removed; live target order/state stays intact.
            document["options"]["targetVisualNames"] = [
                target for target in document["options"]["targetVisualNames"] if target not in missing]
            require(document["options"]["targetVisualNames"], "Repair would leave an empty bookmark target scope")
            repairs.append({"bookmark": name, "removedMissingTargets": sorted(missing)})
        for page, section in state.get('sections', {}).items():
            refs = set(section.get('visualContainers', {})) | set(section.get('visualContainerGroups', {}))
            require(refs <= visual_ids[page], f"Bookmark state targets missing visuals: {name}")
        retained_payload[path] = encode(document) if missing else payload[path]
    # Copy definitions/resources only. Local caches, settings and model files never enter outputs.
    output = {p: b for p, b in payload.items()
              if p.startswith(("StaticResources/", "CustomVisuals/"))
              or p in {"definition/report.json", "definition/version.json"}}
    output.update(page_payload)
    for page in support:
        path = f"definition/pages/{page}/page.json"
        document = decode(output[path])
        document["visibility"] = "HiddenInViewMode"
        output[path] = encode(document)
    output.update(retained_payload)
    if "definition/bookmarks/bookmarks.json" in payload:
        output["definition/bookmarks/bookmarks.json"] = encode(meta)
    page_meta["pageOrder"] = spec["pages"] + support
    page_meta["activePageName"] = spec["pages"][0]
    output["definition/pages/pages.json"] = encode(page_meta)
    report = decode(payload["definition/report.json"])
    for annotation in report.get("annotations", []):
        if annotation.get("name") == "defaultPage":
            annotation["value"] = spec["pages"][0]
    output["definition/report.json"] = encode(report)
    # Reject unsupported report-level page dependencies rather than copying dangling IDs.
    p, b, g = references(report)
    require(p <= selected and b <= retained and g <= groups.keys(), "Unresolved report-level dependency")
    platform = decode(payload[".platform"])
    platform["metadata"]["displayName"] = spec["displayName"]
    platform["metadata"]["description"] = "Generated thematic report; maintain pages in the complete master report."
    platform["config"]["logicalId"] = str(uuid.uuid5(uuid.UUID(platform["config"]["logicalId"]), theme))
    output[".platform"] = encode(platform)
    binding = decode(payload["definition.pbir"])
    reference = binding["datasetReference"]
    model_path = None
    if semantic_model_id:
        model_id = str(uuid.UUID(semantic_model_id))
        binding["datasetReference"] = {"byConnection": {"connectionString": f"semanticmodelid={model_id}"}}
        binding_mode = "shared-service-model-reference-only"
    elif "byPath" in reference:
        require(set(reference) == {"byPath"}, "Ambiguous master model binding")
        model_path = checked_path(source / reference["byPath"]["path"])
        require(model_path.is_dir() and model_path.name.endswith('.SemanticModel'), "Missing shared local model")
        try:
            relative = os.path.relpath(model_path, report_dir).replace(os.sep, '/')
        except ValueError as exc:
            raise GenerationError("Local model and output must share a drive; use an explicit shared model ID") from exc
        binding["datasetReference"]["byPath"]["path"] = relative
        binding_mode = "shared-local-model-reference-only"
    else:
        require(set(reference) == {"byConnection"}, "Unsupported master model binding")
        binding_mode = "shared-connection-preserved"
    output["definition.pbir"] = encode(binding)
    stats = {"generatorVersion": VERSION, "theme": theme, "order": spec["order"],
             "displayName": spec["displayName"], "pages": spec["pages"],
             "supportingPages": support, "bookmarkRepairs": sorted(repairs, key=lambda r: r['bookmark']),
             "visuals": sum(p.endswith('/visual.json') for p in output),
             "bookmarks": len(retained), "syncGroups": sorted({v['groupName'] for data in page_payload.values()
                 for k, v in walk(decode(data)) if k == 'syncGroup' and isinstance(v, dict)}),
             "modelBinding": binding_mode, "masterFiles": fingerprints(original_payload),
             "generatedReportFiles": fingerprints(output), "sourcePreservation": "not-yet-rechecked",
             "publicationPerformed": False, "desktopRenderValidated": False}
    if navigation_changes:
        stats['navigation'] = navigation_changes
    return output, stats, model_path


def generate(source, manifest, theme, output_root, *, write=False, semantic_model_id=None):
    source = checked_path(source)
    output_root = checked_path(output_root)
    require(not output_root.exists(), "Output already exists; choose a new version folder")
    require(not under(output_root, source.parent) and not under(source.parent, output_root),
            "Output must be separate from the master project")
    captured = snapshot(source)
    manifest_bytes = Path(manifest).read_bytes()
    config = read_manifest(manifest_bytes, set(decode(captured['definition/pages/pages.json'])['pageOrder']))
    specs = sorted((spec for spec in config['reports'] if theme == 'all' or spec['key'] == theme),
                   key=lambda spec: spec['order'])
    require(bool(specs), f"Unknown theme: {theme}")
    root_projects = config.get('outputLayout') == 'root-projects'
    require(theme != 'all' or root_projects, "All themes require root-projects layout")
    plans = []
    for spec in specs:
        name = report_file_name(spec)
        report_relative = f"{spec['key']}/{name}.Report" if root_projects else f"{name}.Report"
        report_dir = output_root / report_relative
        receipt_relative = (f"receipts/{spec['key']}-generation-receipt.json.txt" if root_projects
                            else 'generation-receipt.json.txt')
        output, receipt, model = build(source, captured, manifest_bytes, spec['key'], report_dir, semantic_model_id)
        if model:
            require(not under(output_root, model) and not under(model, output_root), "Output overlaps shared model")
        receipt.update(manifestSHA256=digest(manifest_bytes), sourceReport=str(source),
                       outputRoot=str(output_root), projectFile=f"{name}.pbip", reportFolder=report_relative,
                       receiptFile=receipt_relative, written=False)
        plans.append((report_relative, output, receipt))
    require(fingerprints(snapshot(source)) == fingerprints(captured), "Master changed during planning")
    require(Path(manifest).read_bytes() == manifest_bytes, "Manifest changed during planning")
    for _, _, receipt in plans:
        receipt["sourcePreservation"] = "verified-byte-identical"
    result = (plans[0][2] if theme != 'all' else
              {'generatorVersion': VERSION, 'outputLayout': 'root-projects', 'outputRoot': str(output_root),
               'reports': [receipt for _, _, receipt in plans], 'written': False,
               'publicationPerformed': False, 'sourcePreservation': 'verified-byte-identical'})
    if not write:
        return result
    require(output_root.parent.is_dir(), "Output parent must already exist")
    stage = output_root.with_name(output_root.name + '.staging-' + uuid.uuid4().hex)
    stage.mkdir()
    # Failed staging folders are retained for inspection; existing outputs are never removed.
    for report_relative, output, receipt in plans:
        staged_report = stage / report_relative
        for path, data in output.items():
            target = staged_report / path
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open('xb') as handle:
                handle.write(data)
        pbip = {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
                "version": "1.0", "artifacts": [{"report": {"path": report_relative}}],
                "settings": {"enableAutoRecovery": True}}
        (stage / receipt['projectFile']).write_bytes(encode(pbip))
        require(fingerprints(snapshot(staged_report)) == receipt["generatedReportFiles"], "Staged output read-back failed")
        receipt["written"] = True
        receipt_path = stage / receipt['receiptFile']
        receipt_path.parent.mkdir(parents=True, exist_ok=True)
        receipt_path.write_bytes(encode(receipt))
    require(fingerprints(snapshot(source)) == fingerprints(captured), "Master changed during generation")
    require(Path(manifest).read_bytes() == manifest_bytes, "Manifest changed during generation")
    result['written'] = True
    if theme == 'all':
        (stage / 'generation-manifest.json.txt').write_bytes(encode(result))
    require(not output_root.exists(), "Output appeared during generation")
    # rename, not replace: on Windows an existing destination is never overwritten.
    stage.rename(output_root)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source-report', required=True, type=Path)
    parser.add_argument('--manifest', type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument('--theme', default='licensing', help='Theme key or all for a root-projects bundle')
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--semantic-model-id', help='Explicit shared Service model ID; performs no network operation')
    parser.add_argument('--write', action='store_true', help='Create a new output folder; default is plan only')
    args = parser.parse_args()
    try:
        receipt = generate(args.source_report, args.manifest, args.theme, args.output,
                           write=args.write, semantic_model_id=args.semantic_model_id)
    except (GenerationError, OSError, KeyError, TypeError, json.JSONDecodeError, ValueError) as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}), file=sys.stderr)
        return 1
    print(json.dumps(receipt, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
