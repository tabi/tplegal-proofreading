"""Shared minidom XML helpers for DOCX processing.

Common utilities for traversing and querying OOXML DOM trees parsed with
defusedxml.minidom. Used by merge_runs and verify_docx.
"""


def match_local(name: str, tag: str) -> bool:
    """Check if a node name matches a tag, ignoring namespace prefix."""
    return name == tag or name.endswith(f":{tag}")


def find_elements(root, tag: str) -> list:
    """Find all descendant elements matching a local tag name."""
    results = []

    def traverse(node):
        if node.nodeType == node.ELEMENT_NODE:
            name = node.localName or node.tagName
            if match_local(name, tag):
                results.append(node)
            for child in node.childNodes:
                traverse(child)

    traverse(root)
    return results


def extract_paragraph_text(p_elem) -> str:
    """Extract all text from <w:t> elements under a paragraph."""
    texts = []
    for t_elem in find_elements(p_elem, "t"):
        if t_elem.firstChild and hasattr(t_elem.firstChild, "data"):
            texts.append(t_elem.firstChild.data)
    return "".join(texts)
