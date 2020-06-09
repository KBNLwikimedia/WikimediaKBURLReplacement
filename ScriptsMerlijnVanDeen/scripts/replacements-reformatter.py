#https://web.archive.org/web/20200522204706/https://merlijn.vandeen.nl/2015/kb-replace-dead-links.html

f = open('replacements.txt')
out1 = open('replacements_kb_kb.txt', 'w')
out2 = open('replacements_kb_other.txt', 'w')

replacements = set()

for line in f:
    old,new = line.strip().replace('\r', '').split("\t")
    if (old,new) in replacements:
        continue
    replacements.add((old,new))
    print(replacements)

    new = new.replace("https://", "http://")
    print(new)
    if new.startswith("http://www.kb.nl"):
        new = new[11:]
        old = old[11:]
        out1.write(old + "\n" + new + "\n")
        print("aaa",old,new)
    else:pass
     #   old = old.split("www.")[1]
     #   old = r"(https?:)?(//)?(www\.)?" + old
     #   out2.write(old + "\n" + new + "\n")
     #   print("bbb", old, new)