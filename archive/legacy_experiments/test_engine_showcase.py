import os
from pptx import Presentation
from pptx.util import Inches
from ppt_engine import Colors, VisionQA
from ppt_engine.blueprints.timeline import build_callout_timeline_slide, build_chevron_timeline_slide
from ppt_engine.blueprints.process import build_two_row_process_slide

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 1. Timeline with Seamless Callouts
s1_steps = [
    ('Begin', 'Start the project and set the main direction.', Colors.RED, '01'),
    ('Plan', 'Organize tasks and prepare the workflow for the year.', Colors.BLUE, '02'),
    ('Execute', 'Move forward with key actions and major activities.', Colors.GREEN, '03'),
    ('Conclude', 'Wrap up tasks and complete the overall goal.', Colors.ORANGE, '04'),
]
build_callout_timeline_slide(prs, 'Timeline Slide (Seamless Vector Callouts)', s1_steps)

# 2. Alternating Chevron Timeline
s2_steps = [
    ('Start', 'Begin the initiative and set the main direction.', Colors.RED, '01'),
    ('Plan', 'Define yearly targets and organize the required tasks.', Colors.BLUE, '02'),
    ('Build', 'Develop core activities and push the project forward.', Colors.GREEN, '03'),
    ('Improve', 'Review outcomes and enhance areas needing attention.', Colors.ORANGE, '04'),
    ('Strengthen', 'Expand progress and reinforce successful actions.', Colors.TEAL, '05'),
]
build_chevron_timeline_slide(prs, '5-Step Chevron Timeline (Clean Vectors)', s2_steps)

# 3. Two-Row Two-Tone Process Matrix
r1 = [
    ('Market\nAnalysis', '#00A3C4', '#007A93', '01'),
    ('Audience\nIdentification', '#2B6CB0', '#1A4971', '02'),
    ('Lead\nGeneration', '#805AD5', '#553C9A', '03'),
    ('Sales\nInsights', '#E53E3E', '#9B2C2C', '04'),
    ('Pitch\nDevelopment', '#DD6B20', '#9C4221', '05'),
]
r2 = [
    ('Sales\nOutreach', '#C05621', '#7B341E', '06'),
    ('Lead\nNurturing', '#E53E3E', '#9B2C2C', '07'),
    ('Prospect\nEvaluation', '#D69E2E', '#975A16', '08'),
    ('Objection\nHandling', '#2B6CB0', '#1A4971', '09'),
    ('Deal\nClosure', '#00A3C4', '#007A93', '10'),
]
build_two_row_process_slide(prs, '10-Step Business Process (Two-Tone Blocks)', r1, r2)

out_showcase = os.path.abspath('output/engine_showcase.pptx')
prs.save(out_showcase)
print('Saved engine showcase to:', out_showcase)

# Export previews
exported = VisionQA.export_slides_to_png(out_showcase, output_dir='output/engine_showcase_previews')
for exp in exported:
    print('Exported preview:', exp)
