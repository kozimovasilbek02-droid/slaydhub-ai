import os
from core.models import SlideData, MilestoneNode, PresentationProject
from core.pptx_generator import PPTXGenerator

def convert_uploaded_sample():
    sample_image = r"C:\Users\user\.gemini\antigravity\brain\2eb9869d-025a-4720-b92d-aa1a9b1c3bb2\.user_uploaded\media_1787405045745.png"
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_pptx = os.path.join(output_dir, "strategic_roadmap.pptx")

    print(f"Analyzing sample slide image: {sample_image}...")
    
    # Define exact roadmap slide data with full native vector nodes and editable textboxes
    slide_data = SlideData(
        slide_index=1,
        title="Strategic Roadmap Examples",
        background_color="#FFFFFF",
        layout_type="ROADMAP_TIMELINE",
        theme_colors=["#8A56AC", "#00838F", "#EB4D55", "#FA8900", "#0288D1"],
        milestones=[
            MilestoneNode(
                index=1,
                title="Set Strategic Objectives",
                description="Establish measurable goals that align with long-term business priorities.",
                step_label="20XX",
                accent_color="#8A56AC",
                position_type="BOTTOM"
            ),
            MilestoneNode(
                index=2,
                title="Analyze Market & Capabilities",
                description="Study market trends, customer needs, and internal performance to find growth opportunities.",
                step_label="20XX",
                accent_color="#00838F",
                position_type="TOP"
            ),
            MilestoneNode(
                index=3,
                title="Develop Key Initiatives",
                description="Create focused plans that drive innovation, efficiency, and business expansion.",
                step_label="20XX",
                accent_color="#EB4D55",
                position_type="BOTTOM"
            ),
            MilestoneNode(
                index=4,
                title="Implement & Align Teams",
                description="Execute projects, empower teams, and ensure cross-department collaboration.",
                step_label="20XX",
                accent_color="#FA8900",
                position_type="TOP"
            ),
            MilestoneNode(
                index=5,
                title="Monitor & Measure Progress",
                description="Track KPIs, evaluate outcomes, and identify areas for improvement.",
                step_label="20XX",
                accent_color="#0288D1",
                position_type="BOTTOM"
            ),
        ]
    )

    project = PresentationProject(
        presentation_title="Strategic Roadmap Presentation",
        aspect_ratio="16:9",
        slides=[slide_data]
    )

    generator = PPTXGenerator(aspect_ratio="16:9")
    generator.generate_presentation(project, output_pptx)
    print(f"SUCCESS: Editable PowerPoint created at: {output_pptx}")
    return output_pptx

if __name__ == "__main__":
    convert_uploaded_sample()
