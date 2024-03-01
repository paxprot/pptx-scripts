import os
import sys
from pptx import Presentation
from pptx.util import Inches
from lxml import etree


def insert_audio_to_pptx(pptx_file):
    # Load the presentation
    presentation = Presentation(pptx_file)

    # Get the folder name for notes (e.g., "test_notes" for "test.pptx")
    audio_folder = os.path.join(os.path.dirname(pptx_file), "audio")

    # Iterate through slides
    for slide_number, slide in enumerate(presentation.slides, start=1):
        audio_file_path = os.path.join(audio_folder, f"Slide {slide_number}.mp3")

        # Check if audio file exists
        if os.path.exists(audio_file_path):
            audio_shape = slide.shapes.add_movie(
                audio_file_path,
                left=Inches(-1.5),
                top=0,
                width=Inches(1.5),
                height=Inches(1.5),
                poster_frame_image=None,
                mime_type='audio/mp3'
            )

            # Get this workaround to generate autoplay audio from:
            # https://github.com/scanny/python-pptx/issues/427#issuecomment-466217611
            tree = audio_shape._element.getparent().getparent().getnext().getnext()
            timing = [el for el in tree.iterdescendants() if etree.QName(el).localname == 'cond'][0]
            timing.set('delay', '0')

    # Save the modified presentation
    output_pptx_file = os.path.splitext(pptx_file)[0] + "_with_audio.pptx"
    presentation.save(output_pptx_file)

    print(f"Audio files inserted into {output_pptx_file}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python insert_audio.py <input_presentation.pptx>")
        sys.exit(1)

    input_pptx_file = sys.argv[1]
    insert_audio_to_pptx(input_pptx_file)
