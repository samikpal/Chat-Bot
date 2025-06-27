"""Utility to extract hierarchical text chunks from a DOCX file.

The script walks the document once, tracking heading levels so that each
paragraph is associated with its full heading lineage. Chunks begin at
heading paragraphs and include the following normal paragraphs until the
next heading is encountered.

Example usage:
    python parse_docx.py path/to/file.docx
"""

import sys
import uuid
from typing import List, Dict

from docx import Document
from docx.oxml.ns import qn

# Namespaces required for parsing attributes that python-docx does not expose
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

def _paragraph_text(p):
    """Return the concatenated text of a ``w:p`` element including hyperlinks."""
    texts = []
    for t in p.xpath('.//w:t'):
        texts.append(t.text or '')
    return ''.join(texts).strip()

def _paragraph_style(p):
    val = p.xpath('./w:pPr/w:pStyle/@w:val')
    return val[0] if val else None

def _para_id(p):
    return p.get(qn('w14:paraId'))

def _bookmarks(p):
    return [b.get(qn('w:name')) for b in p.xpath('.//w:bookmarkStart')]

def tokenize(text: str) -> int:
    """Very naive token counter."""
    return len(text.split())


def parse_docx(path: str) -> List[Dict]:
    doc = Document(path)
    body = doc._element.body

    chunks = []
    heading_stack: List[str] = []

    current_texts: List[str] = []
    current_para_ids: List[str] = []
    current_bookmarks: List[str] = []
    section = 1

    def finalize():
        if not current_texts:
            return
        chunk_text = "\n\n".join(current_texts)
        chunks.append({
            "chunk_id": uuid.uuid4().hex[:8],
            "text": chunk_text,
            "headings": heading_stack.copy(),
            "para_ids": current_para_ids.copy(),
            "bookmarks": list({b for b in current_bookmarks if b}),
            "section": section,
            "token_count": tokenize(chunk_text),
        })
        current_texts.clear()
        current_para_ids.clear()
        current_bookmarks.clear()

    for child in body.iterchildren():
        tag = child.tag
        if tag.endswith('sectPr'):
            # new section: finalize any open chunk
            finalize()
            section += 1
            continue
        if not tag.endswith('p'):
            continue

        style = _paragraph_style(child)
        text = _paragraph_text(child)
        pid = _para_id(child)
        bookmarks = _bookmarks(child)

        if style and (style.startswith('Heading') or style == 'Title'):
            finalize()
            # determine level
            if style.startswith('Heading'):
                lvl = int(style.replace('Heading', ''))
            else:
                lvl = 1
            # maintain heading stack
            while len(heading_stack) >= lvl:
                heading_stack.pop()
            heading_stack.append(text)
            # start new chunk with heading paragraph
            current_texts.append(text)
            if pid:
                current_para_ids.append(pid)
            current_bookmarks.extend(bookmarks)
        else:
            current_texts.append(text)
            if pid:
                current_para_ids.append(pid)
            current_bookmarks.extend(bookmarks)

    finalize()
    return chunks


def main(argv=None):
    argv = argv or sys.argv[1:]
    if not argv:
        print("Usage: python parse_docx.py <docx-file>")
        return
    chunks = parse_docx(argv[0])
    for chunk in chunks:
        print("-" * 40)
        print(chunk)

if __name__ == "__main__":
    main()
