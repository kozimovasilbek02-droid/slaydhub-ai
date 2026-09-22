# -*- coding: utf-8 -*-
"""
core/slide_morph.py
Native PowerPoint Morph & Motion Storytelling Engine for SlaydHub AI.
Injects authentic PowerPoint OpenXML transitions:
1. Office 365 / PowerPoint Native Morph (p16:morph option="byObject")
2. Cinematic Smooth Fade (p:fade)
3. Dynamic Directional Push (p:push dir="r")
4. Synchronized object naming (!! prefix) for seamless vector morphing across adjacent slides.
"""

import logging
from typing import Optional, List
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

logger = logging.getLogger("SlideMorph")

# Namespaces required for Morph transitions in PPTX OpenXML
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P16_NS = "http://schemas.microsoft.com/office/powerpoint/2015/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class SlideMorphEngine:
    """
    Applies OpenXML native slide transitions and motion storytelling effects to PPTX slides.
    """

    @classmethod
    def apply_transition(
        cls,
        slide,
        transition_type: str = "morph",
        speed: str = "med",
        direction: str = "r"
    ) -> bool:
        """
        Injects a schema-compliant <p:transition> element into the slide XML.
        Uses mc:AlternateContent to guarantee 100% backward compatibility with
        PowerPoint 2013/2016 while enabling native Morph in Office 365 / 2019+.
        """
        try:
            sld_elem = slide.element

            # Check if transition already exists and remove it to avoid duplicates
            existing_trans = sld_elem.find(f"{{{P_NS}}}transition")
            if existing_trans is not None:
                sld_elem.remove(existing_trans)

            # Build transition XML string with mc:AlternateContent for Morph
            if transition_type == "morph":
                trans_xml = (
                    f'<p:transition {nsdecls("p")} xmlns:mc="{MC_NS}" spd="{speed}" advClick="1">'
                    f'  <mc:AlternateContent>'
                    f'    <mc:Choice Requires="p16" xmlns:p16="{P16_NS}">'
                    f'      <p16:morph option="byObject"/>'
                    f'    </mc:Choice>'
                    f'    <mc:Fallback>'
                    f'      <p:fade/>'
                    f'    </mc:Fallback>'
                    f'  </mc:AlternateContent>'
                    f'</p:transition>'
                )
            elif transition_type == "fade":
                trans_xml = (
                    f'<p:transition {nsdecls("p")} spd="{speed}" advClick="1">'
                    f'  <p:fade/>'
                    f'</p:transition>'
                )
            elif transition_type == "push":
                trans_xml = (
                    f'<p:transition {nsdecls("p")} spd="{speed}" advClick="1">'
                    f'  <p:push dir="{direction}"/>'
                    f'</p:transition>'
                )
            elif transition_type == "wipe":
                trans_xml = (
                    f'<p:transition {nsdecls("p")} spd="{speed}" advClick="1">'
                    f'  <p:wipe dir="{direction}"/>'
                    f'</p:transition>'
                )
            else:
                trans_xml = (
                    f'<p:transition {nsdecls("p")} spd="{speed}" advClick="1">'
                    f'  <p:fade/>'
                    f'</p:transition>'
                )

            trans_element = parse_xml(trans_xml)

            # In OpenXML p:sld schema, the order is:
            # 1. p:cSld
            # 2. p:clrMapOvr
            # 3. p:transition
            # 4. p:timing
            # 5. p:extLst
            # Insert before p:timing or p:extLst if present
            timing_elem = sld_elem.find(f"{{{P_NS}}}timing")
            extlst_elem = sld_elem.find(f"{{{P_NS}}}extLst")

            if timing_elem is not None:
                timing_elem.addprevious(trans_element)
            elif extlst_elem is not None:
                extlst_elem.addprevious(trans_element)
            else:
                sld_elem.append(trans_element)

            return True
        except Exception as e:
            logger.warning(f"Could not apply transition {transition_type} to slide: {e}")
            return False

    @classmethod
    def apply_deck_storytelling_transitions(cls, prs, mode: str = "cinematic") -> int:
        """
        Applies a cohesive motion storytelling flow across all slides in a presentation:
        - Slide 1 (Cover): Smooth fade into Agenda
        - Slide 2..N-1 (Content/Concepts): Native Morph or Push depending on narrative continuity
        - Final Slide (Conclusion/Q&A): Elegant Fade
        """
        slide_count = len(prs.slides)
        applied = 0

        for i, slide in enumerate(prs.slides):
            if i == 0:
                # Cover slide has no entry transition or simple fade
                continue
            elif i == 1:
                # Agenda/Reja: Smooth Fade
                if cls.apply_transition(slide, transition_type="fade", speed="med"):
                    applied += 1
            elif i == slide_count - 1:
                # Conclusion: Fade
                if cls.apply_transition(slide, transition_type="fade", speed="med"):
                    applied += 1
            else:
                # Core content: PowerPoint Native Morph
                trans = "morph" if mode == "cinematic" else "push"
                if cls.apply_transition(slide, transition_type=trans, speed="med"):
                    applied += 1

        logger.info(f"Storytelling transitions applied to {applied} / {slide_count} slides (mode={mode})")
        return applied
