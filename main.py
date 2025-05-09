#USER DEFINED CONSTANTS
CARD_IMAGE_DIRECTORY = "images"
PROJECTIVE_PLANE_SHEET = "projective_plane.xlsx"
SPOT_IT_GAME_SLIDESHOW = "spot_it_game.pptx"


#DO NOT MODIFY CODE BELOW HERE

#Imports and set up
from pptx import Presentation
from pptx.util import Inches
import random
from random import shuffle
import pandas as pd
import os
from sys import exit

#Loads in the images from the directory
images_path = os.path.join(os.path.curdir, CARD_IMAGE_DIRECTORY)
images_names = os.listdir(images_path)
images_path_names = [os.path.join(images_path, image_name) for image_name in images_names]

#Checks for right file extensions and count
valid_extensions = ["png", "jpg", "jpeg", "gif"]
for image in images_path_names:
  if not os.path.isfile(image):
    print("Please do not place any subfolders in the images folder")
    exit()
  extension = image.split(os.extsep)[-1]
  if extension not in valid_extensions:
    print(f"File type '{extension}' is not supported. Please use one of the following: {', '.join(valid_extensions)}")
    exit()

if len(images_names) != 57:
  print(f"You have placed '{len(images_names)}' files in images. Please instead place 57.")
  exit()

#Loads in the card indicies from the projective plane
cards_df = pd.read_excel(PROJECTIVE_PLANE_SHEET)
cards = cards_df.values.tolist()
print(cards_df)

#Grid position for slide
grid_positions = [(0, 0), (1, 0), (2, 0), (0, 1), (2, 1), (0, 2), (1, 2), (2, 2)]  # 8 out of 9 grid spots

#Slide and image dimensions
slide_width = Inches(10)
slide_height = Inches(7.5)
img_width = Inches(2.5)
img_height = Inches(2.5)
x_margin = Inches(0.75)
y_margin = Inches(0.0)
x_spacing = Inches(3)
y_spacing = Inches(2.5)

#Create PowerPoint
pres = Presentation()
blank_layout = pres.slide_layouts[6]

for card in cards:
  shuffle(grid_positions)
  slide = pres.slides.add_slide(blank_layout)

  for i, symbol_num in enumerate(card[1:]):
    col, row = grid_positions[i]
    x = x_margin + col * x_spacing
    y = y_margin + row * y_spacing

    image_path = images_path_names[symbol_num-1]
    print(image_path)

    if os.path.exists(image_path):
      img = slide.shapes.add_picture(image_path, x, y, img_width, img_height)
      img.rotation = random.randint(0, 360)
    else:
      print(f"Missing image: {image_path}")

#Save
pres.save(SPOT_IT_GAME_SLIDESHOW)
print(f"Presentation successfully saved as {SPOT_IT_GAME_SLIDESHOW}")