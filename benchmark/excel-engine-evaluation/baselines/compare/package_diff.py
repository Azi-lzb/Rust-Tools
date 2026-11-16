"""Byte-level uncompressed part hashes, plus separate XML canonical comparison."""
import hashlib
import xml.etree.ElementTree as ET
import zipfile


def parts(path):
    with zipfile.ZipFile(path) as z:
        names=z.namelist()
        if len(names)!=len(set(names)):
            raise AssertionError('duplicate ZIP entries')
        return {name:z.read(name) for name in names}


def package_diff(source, target):
    a,b=parts(source),parts(target)
    changed=[]
    for name in sorted(a.keys() & b.keys()):
        if a[name] != b[name]:
            item=dict(part=name,source_sha256=hashlib.sha256(a[name]).hexdigest(),
                      target_sha256=hashlib.sha256(b[name]).hexdigest())
            if name.endswith(('.xml','.rels')):
                try: item['canonical_xml_equal']=ET.canonicalize(a[name])==ET.canonicalize(b[name])
                except (ET.ParseError,UnicodeError): item['canonical_xml_equal']=False
            changed.append(item)
    return dict(added=sorted(b.keys()-a.keys()),removed=sorted(a.keys()-b.keys()),changed=changed,
                unchanged=sorted(n for n in a.keys() & b.keys() if a[n]==b[n]))
