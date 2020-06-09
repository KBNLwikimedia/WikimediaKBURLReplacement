# https://web.archive.org/web/20200522204706/https://merlijn.vandeen.nl/2015/kb-replace-dead-links.html
# hervorm 'urls.txt' tot 'nlpages.txt' en 'otherpages.txt'

import re

matcher = re.compile(r".*?//(\w+)\..*?/wiki/(.*)")

nlout = open('nlpages.txt', 'w')
oout = open('otherpages.txt', 'w')

for line in open('urls.txt'):
    line = line.strip()
    match = matcher.match(line)
    if not match:
        print("Failed to match", line)
    else:
        lang, pagename = match.groups()
        line = "[[%s:%s]]\n" % (lang, pagename)
        if lang == "nl":
            nlout.write(line)
        else:
            oout.write(line)