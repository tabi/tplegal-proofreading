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



def extract_paragraph_text(p_elem, mode: str = "visible") -> str:
    """Extract text from a paragraph with tracked-change awareness.

    Args:
        p_elem: A minidom paragraph element.
        mode: 'visible' — text as Word displays it (skip w:delText, include w:ins/w:t).
              'original' — text before corrections (include w:delText, skip w:ins/w:t).
    """
    texts = []

    def _collect(node):
        if node.nodeType == node.ELEMENT_NODE:
            name = node.localName or node.tagName

            if mode == "visible":
                # Skip entire w:del subtrees (deleted text not visible)
                if match_local(name, "del"):
                    return
            elif mode == "original":
                # Skip entire w:ins subtrees (inserted text didn't exist originally)
                if match_local(name, "ins"):
                    return

            # Collect text from w:t and w:delText
            if match_local(name, "t") or match_local(name, "delText"):
                if node.firstChild and hasattr(node.firstChild, "data"):
                    texts.append(node.firstChild.data)
                return  # no need to recurse into text nodes

            for child in node.childNodes:
                _collect(child)

    _collect(p_elem)
    return "".join(texts)
