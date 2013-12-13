import re
from six.moves import zip

def check_tag(root, expected):
    pattern = re.compile(r"{.*}([a-zA-Z]+)")
    for tag, el in zip(expected, root.iter()):
        m = pattern.match(el.tag)
        assert m is not None
        assert m.group(1) == tag, "Expect tag=%s, get %s" % (tag, m.group(1))

