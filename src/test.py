import zipfile
from lxml import etree

def print_tree(elem, depth=0):
    indent = "  " * depth

    # Clean tag (remove namespace)
    tag = elem.tag.split("}")[-1]

    # Attributes
    attrs = " ".join([f'{k.split("}")[-1]}="{v}"' for k, v in elem.attrib.items()])
    if attrs:
        attrs = " " + attrs

    # Text
    text = (elem.text or "").strip()
    if text:
        print(f"{indent}<{tag}{attrs}> {text}")
    else:
        print(f"{indent}<{tag}{attrs}>")

    for child in elem:
        print_tree(child, depth + 1)

    print(f"{indent}</{tag}>")

def dump_docx_dom(docx_path):
    with zipfile.ZipFile(docx_path) as docx:
        xml = docx.read("word/document.xml")

    root = etree.XML(xml)

    print_tree(root)

if __name__ == "__main__":
    dump_docx_dom("data/templates/base_template.docx")
