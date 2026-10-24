"""Associate embedded media with nearby Word paragraphs in XML order."""
import re
from collections import defaultdict

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
CAPTION = re.compile(r'^(?:图\s*[\d一二三四五六七八九十]|figure\s*\d|fig\.\s*\d)', re.I)
RESERVED = re.compile(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$', re.I)


def filename_from_text(text):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', ' ', text)
    name = re.sub(r'\s+', ' ', name).strip(' .')[:100].rstrip(' .')
    if name and RESERVED.fullmatch(name.split('.')[0]):
        name = '_' + name
    return name


def image_contexts(root, mapping, document_name):
    parents = {child: parent for parent in root.iter() for child in parent}
    paragraphs = list(root.iter(W + 'p'))
    groups = defaultdict(list)
    for number, paragraph in enumerate(paragraphs, 1):
        scope = parents.get(paragraph)
        while scope is not None and scope.tag not in (W + 'tc', W + 'txbxContent', W + 'body'):
            scope = parents.get(scope)
        tokens, alternatives = [], {}

        def visit(node):
            # Text boxes contain their own paragraphs; don't duplicate them here.
            if node is not paragraph and node.tag == W + 'p':
                return
            if node.tag in (W + 'del', W + 'moveFrom'):
                return
            target = mapping.get(node.get(R + 'embed') or node.get(R + 'id'), '')
            if target.startswith('word/media/'):
                tokens.append(('image', target))
                container = node
                while container.tag not in (W + 'drawing', W + 'pict') and parents.get(container) is not paragraph and parents.get(container) is not None:
                    container = parents[container]
                descriptions = alternatives.setdefault(target, [])
                for child in container.iter():
                    if child.tag.endswith('}docPr'):
                        for text in (child.get('title'), child.get('descr')):
                            if text and text not in descriptions:
                                descriptions.append(text)
            elif node.tag == W + 't':
                tokens.append(('text', node.text or ''))
            elif node.tag in (W + 'tab', W + 'br', W + 'cr'):
                tokens.append(('text', ' '))
            for child in node:
                visit(child)

        visit(paragraph)
        text = ''.join(value for kind, value in tokens if kind == 'text').strip()
        style = paragraph.find(f'{W}pPr/{W}pStyle')
        record = {'number': number, 'scope': scope, 'tokens': tokens, 'text': text,
                  'has_image': any(kind == 'image' for kind, _ in tokens),
                  'caption': bool(style is not None and re.search(r'caption|题注', style.get(W + 'val', ''), re.I)),
                  'alt': alternatives}
        groups[scope].append(record)

    contexts = defaultdict(list)
    for group in groups.values():
        for index, record in enumerate(group):
            if not record['has_image']:
                continue

            def neighbors(direction):
                found = []
                for distance in range(1, 7):
                    position = index + distance * direction
                    if not 0 <= position < len(group):
                        break
                    neighbor = group[position]
                    previous = group[position - direction]
                    if abs(neighbor['number'] - previous['number']) != 1:
                        break
                    # Don't borrow another image's caption from beyond that image.
                    if neighbor['has_image']:
                        break
                    if neighbor['text']:
                        found.append({'text': neighbor['text'], 'caption': neighbor['caption']})
                        if len(found) == 2:
                            break
                return found

            above, below = neighbors(-1), neighbors(1)
            for token_index, (kind, target) in enumerate(record['tokens']):
                if kind != 'image':
                    continue
                before, after = [], []
                for direction, output in ((-1, before), (1, after)):
                    position = token_index + direction
                    while 0 <= position < len(record['tokens']):
                        token_kind, value = record['tokens'][position]
                        if token_kind == 'image':
                            break
                        output.append(value)
                        position += direction
                context = {
                    'document': document_name, 'paragraph': record['number'],
                    'inline_before': ''.join(reversed(before)).strip(),
                    'inline_after': ''.join(after).strip(),
                    'above': [item['text'] for item in reversed(above)],
                    'below': [item['text'] for item in below],
                    'alt_text': record['alt'].get(target, []), 'candidates': [],
                }
                options = [(context['inline_before'], '同段前文', record['caption']),
                           (context['inline_after'], '同段后文', record['caption'])]
                options += [(item['text'], '图片下方', item['caption']) for item in below]
                options += [(item['text'], '图片上方', item['caption']) for item in above]
                options += [(text, '替代文字', False) for text in context['alt_text']]
                seen = set()
                for text, label, caption in options:
                    name = filename_from_text(text)
                    if not name or name.casefold() in seen:
                        continue
                    seen.add(name.casefold())
                    context['candidates'].append({'text': text, 'name': name, 'source': label,
                                                  'is_caption': caption or bool(CAPTION.match(text.strip()))})
                context['candidates'].sort(key=lambda c: not c['is_caption'])
                if context not in contexts[target]:
                    contexts[target].append(context)
    return contexts
