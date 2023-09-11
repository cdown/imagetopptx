#!/usr/bin/env python

from pptx import Presentation
from pptx.util import Inches
import os
import sys

# TODO: is this always the empty one?
EMPTY_SLIDE_LAYOUT = 5

prs = Presentation()

# TODO: aspect ratio only, but seems to be enough at least for google slides
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)
slide_layout = prs.slide_layouts[EMPTY_SLIDE_LAYOUT]

for image_file in sys.argv[1:]:
    if image_file.endswith(".png"):
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.add_picture(image_file, 0, 0, prs.slide_width, prs.slide_height)

# Save the presentation
prs.save("presentation.pptx")
