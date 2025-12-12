# This is a color palette app that shows complementary, analogous, and triadic color combinations when one color is entered (or generated randomly). It can also provide the user with a poster or presentation template using the color palette.

Created using Python 3.14.0

# Overview of UI:
<img width="1246" height="1261" alt="Screenshot 2025-12-08 095225" src="https://github.com/user-attachments/assets/0411f332-ff15-400d-9ecf-c99d7db08b3a" />

# Functionality:
# Three options for choosing a starting color
<img width="1226" height="111" alt="Screenshot 2025-12-08 095147" src="https://github.com/user-attachments/assets/6e683b04-6810-4df7-a1c3-c84fc3eb1d92" />

1. Press the "Choose a Color" button to choose a color using a classic color picker
<img width="885" height="657" alt="Screenshot 2025-12-08 105251" src="https://github.com/user-attachments/assets/d7ce5df0-cdec-4c0e-b762-d31707375824" />

2. Press the "Open Image" button to upload an image from your computer and color-pick from the image. The color and its RGB values will show up next to the mouse cursor.
- When clicked, the RGB values will also show up in the lower right-hand corner of the main display area and be set as the "starting color".
<img width="407" height="343" alt="Screenshot 2025-12-08 105612" src="https://github.com/user-attachments/assets/722972d9-4432-4939-ba0a-709e00cfbf11" />

3. Press the "Random Color!" button to generate a random starting color.

# Three options for palette types. 

<img width="670" height="670" alt="Linear_RGB_color_wheel" src="https://github.com/user-attachments/assets/57fd5302-a09a-4d97-b5f7-a3426b2337e7" />

1. Complementary will generate a color that is on the opposite side of the color wheel (from the starting color).
2. Analogous will generate two colors that are 30 degrees away from the starting color in either direction on the color wheel.
3. Triadic will generate two colors that are 120 degrees away from the starting color in either direction on the color wheel.

- Select the desired palette type to generate a previous of the 2-3 colors. RGB values will also be generated and displayed below the color blocks.
  - Note that in Release v1.0, the RGB values are not able to be highlighted and copied.
<img width="1217" height="985" alt="Screenshot 2025-12-08 104516" src="https://github.com/user-attachments/assets/ebfef3b6-2435-4394-b895-11e549739b94" />

# PowerPoint Templates
Two options for templates to generate: An academic poster and presentation template. Both are in pptx format, and will save to the folder from which the user launched the program. Saving a new pptx file won't overwrite any existing files.

1. Poster:
<img width="1448" height="968" alt="Screenshot 2025-12-08 110739" src="https://github.com/user-attachments/assets/723dca46-10fe-4d2b-b571-7de14b9b823e" />
- The poster will be generated with the dimensions of 36x24 inches. If the conference requires a larger size, the user can modify the dimensions in PowerPoint by going to Design > Slide Size > Custom Slide Size and adjust the font sizes. Alternatively, some printers allow one to scale up the size of a document when printing. <b>Be sure to save as a PDF for best results!</b>

<br>
2. Presentation:
<img width="1995" height="1007" alt="Screenshot 2025-12-08 104551" src="https://github.com/user-attachments/assets/47076cc5-c211-4a87-9e97-da14cfc2c3e1" />


Press the x in the top right corner to close the program.

(Python logo and color wheel image from Wikipedia.)




