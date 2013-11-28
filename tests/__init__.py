import re
from itertools import izip

def check_tag(root, expected):
    pattern = re.compile(r"{.*}([a-zA-Z]+)")
    for tag, el in izip(expected, root.iter()):
        m = pattern.match(el.tag)
        assert m is not None
        assert m.group(1) == tag, "Expect tag=%s, get %s" % (tag, m.group(1))

