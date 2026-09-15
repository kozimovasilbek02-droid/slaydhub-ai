import os
import argparse
from core.models import PresentationProject
from core.vision_parser import VisionSlideParser
from core.pptx_generator import PPTXGenerator

def main():
    parser = argparse.ArgumentParser(description="Convert Slide Images (.jpg/.png) into 100% Editable PowerPoint Presentation (.pptx)")
    parser.add_argument("input_path", help="Path to an image file or folder containing slide images")
    parser.add_argument("-o", "--output", default="output/converted_presentation.pptx", help="Path to output .pptx file")
    parser.add_argument("--api-key", default=None, help="Google Gemini API key for AI vision parsing")
    parser.add_argument("--ratio", choices=["16:9", "4:3"], default="16:9", help="Presentation aspect ratio")
    
    args = parser.parse_args()

    images = []
    if os.path.isfile(args.input_path):
        images.append(args.input_path)
    elif os.path.isdir(args.input_path):
        for f in sorted(os.listdir(args.input_path)):
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                images.append(os.path.join(args.input_path, f))

    if not images:
        print("No slide images found!")
        return

    print(f"Found {len(images)} slide image(s). Decompiling into editable PPTX...")
    parser_ai = VisionSlideParser(api_key=args.api_key)
    slides = []
    
    for i, img_path in enumerate(images):
        print(f"Processing slide [{i+1}/{len(images)}]: {img_path}...")
        slide_data = parser_ai.parse_image(img_path)
        slide_data.slide_index = i + 1
        slides.append(slide_data)

    project = PresentationProject(
        presentation_title="Converted Presentation",
        aspect_ratio=args.ratio,
        slides=slides
    )

    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    generator = PPTXGenerator(aspect_ratio=args.ratio)
    generator.generate_presentation(project, args.output)
    print(f"Successfully generated: {args.output}")

if __name__ == "__main__":
    main()
