import argparse
import lxml
from lxml import etree, objectify
import os.path
import html

# Check if the input file is valid
def is_valid_file(parser, arg):
  if not os.path.exists(arg):
    parser.error("The file %s does not exist!" % arg)
  else:
    return open(arg, 'r')  # return an open file handle

# Define the program options
parser = argparse.ArgumentParser(description='While using the Mammoth docx converter, add options to preserve source class names and formatting information as attributes.')
parser.add_argument("-i", dest="filename", required=True,
                    help="The name of the file to read.", metavar="FILE",
                    type=lambda x: is_valid_file(parser, x))

args = parser.parse_args()

root = etree.parse(args.filename.name)

sectionstarts = ["hsectitlepage",
                 "hsecchapter",
                 "hsechalftitlepage",
                 "hseccopyright-page",
                 "hsecdedication",
                 "hsecepigraph",
                 "hsecforeword",
                 "hsecpreface",
                 "hsectoc",
                 "hsecintroduction",
                 "hsecpart",
                 "hsecinterlude",
                 "hsecappendix",
                 "hseccolophon",
                 "hsecacknowledgments",
                 "hsecafterword",
                 "hsecconclusion",
                 "hsecglossary",
                 "hsecbibliography",
                 "hsecabout-the-author",
                 "hsecindex",
                 "hsecendnotes"]

# order of operations is important;
# wpr processing should happen first
for wpr in root.xpath(".//*[re:test(@class, 'hwpr[a-zA-Z0-9-]+start')]", namespaces={"re": "http://exslt.org/regular-expressions"}):
  mystart = wpr.get("class")
  myend = mystart.replace("start", "end")
  wprclass = mystart.replace("hwpr", "").replace("start", "")
  nextblock = wpr.getnext()
  while nextblock is not None and nextblock.get("class") != myend:
    currclass = nextblock.get("class")
    if "hwpr" not in currclass:
      currclass = currclass.split(" ")
      if len(currclass) > 1:
        origblk = currclass.pop()
        s = " "
        newclass = s.join(currclass) + " " + wprclass + " " + origblk
      else:
        newclass = wprclass + " " + currclass[0]
      nextblock.set("class", newclass)
      nextblock = nextblock.getnext()
    else:
      nextblock = nextblock.getnext()
  wpr.getparent().remove(wpr)

for wpr in root.xpath(".//*[re:test(@class, 'hwpr[a-zA-Z0-9-]+end')]", namespaces={"re": "http://exslt.org/regular-expressions"}):
  wpr.getparent().remove(wpr)

for sectstart in sectionstarts:
  for para in root.findall(".//*[@class='" + sectstart + "']"):
    nextblock = para.getnext()
    while nextblock is not None and nextblock.get("class") not in sectionstarts:
      currclass = nextblock.get("class")
      newclass = sectstart + " " + currclass
      nextblock.set("class", newclass)
      nextblock = nextblock.getnext()

for sectstart in sectionstarts:
  for para in root.findall(".//*[@class='" + sectstart + "']"):
    para.getparent().remove(para)

def sanitizeHTML(myroot):
  newHTML = etree.tostring(myroot, encoding="UTF-8", standalone=True, xml_declaration=True)
  return newHTML

myhtml = sanitizeHTML(root)

outputfilename = args.filename.name.split(".")[0] + "_flattened.html"
output = open(outputfilename, 'wb')
output.write(myhtml)
output.close()

# h_wpr_ordered-list_a_start": "h_wpr_ordered-list_a_end"