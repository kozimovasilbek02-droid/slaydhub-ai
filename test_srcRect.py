from pptx.oxml.xmlchemy import OxmlElement

srcRect = OxmlElement('a:srcRect')
srcRect.set('l', '10000')
print(srcRect.xml)
