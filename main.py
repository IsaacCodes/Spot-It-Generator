from pptx import Presentation
from pptx.util import Inches
import pandas as pd
import os

image_path = os.path.join(os.path.curdir, "images")

for file in os.listdir(image_path):
  print(file)


os.abort()
print("yo")
# Constants
IMAGE_FOLDER = "images"
GRID_POSITIONS = [(0, 0), (1, 0), (2, 0),
            (0, 1), (1, 1), (2, 1),
            (0, 2), (1, 2)]  # 8 out of 9 grid spots

# Slide and image dimensions
slide_width = Inches(10)
slide_height = Inches(7.5)
img_width = Inches(2.5)
img_height = Inches(2.5)
x_margin = Inches(0.5)
y_margin = Inches(0.5)
x_spacing = Inches(3)
y_spacing = Inches(2.5)

# Load validated projective plane (replace this with your actual data)
# Example: cards = [[2, 9, 16, 23, 30, 37, 44, 51], ...]
cards_df = pd.read_excel("projective_plane.xlsx")  # <- make sure to match filename
cards = cards_df.iloc[:, 1:].values.tolist()

# Create PowerPoint
prs = Presentation()
blank_layout = prs.slide_layouts[6]

for idx, card in enumerate(cards):
  slide = prs.slides.add_slide(blank_layout)
  
  for i, symbol_num in enumerate(card):
    col, row = GRID_POSITIONS[i]
    x = x_margin + col * x_spacing
    y = y_margin + row * y_spacing
    image_path = os.path.join(IMAGE_FOLDER, f"{IMAGE_PREFIX}{symbol_num}{IMAGE_SUFFIX}")
    if os.path.exists(image_path):
      slide.shapes.add_picture(image_path, x, y, width=img_width, height=img_height)
    else:
      print(f"Missing image: {image_path}")

# Save
prs.save("spot_it_game_cards.pptx")
print("Presentation saved as spot_it_game_cards.pptx")
