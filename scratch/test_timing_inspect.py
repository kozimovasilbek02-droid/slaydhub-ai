from pptx import Presentation

prs = Presentation("output/test_authentic_wingdings_preserved.pptx")
slide = prs.slides[11]
valid_ids = set(slide.element.xpath('.//*[local-name()="cNvPr"]/@id'))
print("Valid IDs:", sorted(list(valid_ids), key=lambda x: int(x) if x.isdigit() else 0))

timings = slide.element.xpath('./*[local-name()="timing"]')
if timings:
    timing = timings[0]
    all_spids = set(timing.xpath('.//*[@spid]/@spid'))
    print("Targeted spids in timing:", all_spids)
    zombies = all_spids - valid_ids
    print("Zombie spids:", zombies)
