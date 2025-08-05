# coding: utf-8
u"""
This code is written very similarly to docx_utils :func:`~docx_utils.flatten.opc_to_flat_opc`
with the base idea coming from Eric White's blog `Transforming Flat OPC Format to Open XML Documents
<https://web.archive.org/web/20130511174646/http://blogs.msdn.com/b/ericwhite/archive/2008/09/29/transforming-flat-opc-format-to-open-xml-documents.aspx/>`
"""
import base64
import collections
import io
import os
import zipfile

from lxml import etree


class ContentTypes(object):
    """
    ContentTypes contained in a "[Content_Types].xml" file.
    """

    CT = "http://schemas.openxmlformats.org/package/2006/content-types"
    NS = {"ct": CT}

    def __init__(self):
        self._defaults = {}
        self._overrides = {}

    def write_xml_data(self):
        content = ('<Types xmlns="{ct}"/>').format(ct=self.CT)
        doc: etree._ElementTree = etree.parse(io.StringIO(content))
        root = doc.getroot()

        for key, value in self._defaults:
            node: etree._Element = etree.SubElement(
                root, "{{{ct}}}Default".format(ct=self.CT), nsmap=self.NS
            )
            node.set("Extension", key)
            node.set("ContentType", value)

        for key, value in self._overrides.items():
            node: etree._Element = etree.SubElement(
                root, "{{{ct}}}Override".format(ct=self.CT), nsmap=self.NS
            )
            node.set("PartName", key)
            node.set("ContentType", value)

        return etree.tostring(
            doc,
            xml_declaration=True,
            encoding="UTF-8",
            pretty_print=False,
            with_tail=False,
            standalone=True,
        )

    def add_content(self, part_name, content_type):
        self._overrides[part_name] = content_type


PackagePart = collections.namedtuple("PackagePart", ["uri", "content_type", "data"])


def find_packages(flat_path):
    """
    Iterate a Flat XML document and yield the package parts.

    :param str flat_path: Microsoft Office document in flat OPC format (.xml) or the contents of the file as a string

    :return: Iterator which yield package parts
    """
    pkg = "http://schemas.microsoft.com/office/2006/xmlPackage"
    ns = {"pkg": pkg}

    if os.path.exists(flat_path):
        with io.open(flat_path, mode="rb") as f:
            content = f.read()
    else:
        content = flat_path.encode("utf-8")

    try:
        tree: etree._ElementTree = etree.fromstring(content)
    except etree.XMLSyntaxError as e:
        if "huge text node" in str(e):
            parser = etree.XMLParser(huge_tree=True)
            tree: etree._ElementTree = etree.fromstring(content, parser=parser)
        else:
            raise e
    
    for part in tree.xpath("//pkg:part", namespaces=ns):
        uri = part.attrib["{{{pkg}}}name".format(pkg=pkg)]
        content_type = part.attrib["{{{pkg}}}contentType".format(pkg=pkg)]
        if content_type.endswith("xml"):
            data = etree.tostring(
                list(list(part)[0])[0],
                xml_declaration=True,
                encoding="UTF-8",
                pretty_print=False,
                with_tail=False,
                standalone=True,
            )
        else:
            chunks = list(part)[0].text
            encoded = list(filter(lambda char: char not in ["\n", "\r"], chunks))
            data = base64.b64decode("".join(encoded).encode())
        yield PackagePart(uri, content_type, data)


def flat_to_opc(src_path, dest_path):
    """
    Convert an flat OPC document into a full OPC zipped file.

    :param str src_path: Microsoft Office document in flat OPC format (.xml) or string containing files contents

    :param str dest_path: Microsoft Office document convert to (.docx, .xlsx, .pptx)
    """
    content_types = ContentTypes()

    with zipfile.ZipFile(dest_path, mode="w") as f:
        for file in find_packages(src_path):
            content_types.add_content(file.uri, file.content_type)
            f.writestr(file.uri[1:], file.data)
        f.writestr("[Content_Types].xml", content_types.write_xml_data())


def flat_to_opc_bytes(src_path) -> bytes:
    """
    Convert an flat OPC document into a full OPC zipped file.

    :param str src_path: Microsoft Office document in flat OPC format (.xml) or string containing files contents

    :return bytes: Microsoft Office document converted to (.docx, .xlsx, .pptx) in binary
    """
    content_types = ContentTypes()

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, mode="w") as f:
        for file in find_packages(src_path):
            content_types.add_content(file.uri, file.content_type)
            f.writestr(file.uri[1:], file.data)
        f.writestr("[Content_Types].xml", content_types.write_xml_data())

    return zip_buffer.getvalue()
