## This is a color palette app that shows complementary, analogous, and triadic color combinations when one color is entered (or generated randomly). It can also provide the user with a poster or presentation template using the color palette.

## Created using Python 3.14.0

# Developer's Guide
## Condensed version of final planning spec:
- User can choose a starting color three different ways: classic color picker, uploading an image file, or generating a random color.
- User can select one of three color palettes (complementary, analogous, triadic) and see a preview of the starting color with the 1-2 other colors in block format.
- User can generate and download an academic poster or presentation template (both pptx) using the color palette.

## Overview of UI:
<img width="1246" height="1261" alt="Screenshot 2025-12-08 095225" src="https://github.com/user-attachments/assets/0411f332-ff15-400d-9ecf-c99d7db08b3a" />

# Functionality:
## Starting Code
The program is created inside of a TkInter interface. First inside the init are the UI elements and global variables (current_rgb, palette_rgb1, palette_grb2). These variables are set to have no values to start, and will be reassigned values by the starting color and palette creator functions.

## Three options for choosing a starting color
<img width="1226" height="111" alt="Screenshot 2025-12-08 095147" src="https://github.com/user-attachments/assets/6e683b04-6810-4df7-a1c3-c84fc3eb1d92" />

1. Press the "Choose a Color" button to choose a color using a classic color picker
- This is handled by the colorpicker() function. It calls the TkInter colorchooser method, which returns a tuple with an RGB and hex value. We only need the RGB, so only take the first element. That RGB value is set as the current_rgb value (starting value) and appears in the terminal and the bottom-right corner of the canvas in the UI. A rectangle of the color is displayed in the UI.
<img width="885" height="657" alt="Screenshot 2025-12-08 105251" src="https://github.com/user-attachments/assets/d7ce5df0-cdec-4c0e-b762-d31707375824" />

3. Press the "Open Image" button to upload an image from your computer and color-pick from the image. The color and its RGB values will show up next to the mouse cursor.
- This is handled by the open_image function, allowing the file to be chosen from a file browser on the user's computer and setting a thumbnail in the canvas.
- Then, the track_cursor function tracks which pixel of the image thumbnail the mouse is hovering over. The x and y coordinates update as the cursor moves. The getpixel(event.x, event.y) method of PIL gets the RGB value of the pixel. Then it needs to be converted to hex for use with the color_label label, which displays the RGB value of the pixel.
- Clicking on a pixel in the image calls the get_color(event) function, which updates current_rgb variable (starting color) and displays it in lower-right corner. Clicking outside of the image throws a exception: "Cursor is outside the image bounds."
<img width="407" height="343" alt="Screenshot 2025-12-08 105612" src="https://github.com/user-attachments/assets/722972d9-4432-4939-ba0a-709e00cfbf11" />

3. Press the "Random Color!" button to generate a random starting color.
- This is handled by the random_color() function. Three random values between 0 and 255 (inclusive) are generated for the red, green, and blue channels. A rectangle of the color is also generated and displayed in the UI. The RGB value is set as the current_rgb variable value and displayed in the lower-right corner.

## Three options for palette types. 

<img width="670" height="670" alt="Linear_RGB_color_wheel" src="https://github.com/user-attachments/assets/57fd5302-a09a-4d97-b5f7-a3426b2337e7" />

1. Complementary will generate a color that is on the opposite side of the color wheel (from the starting color).
- This is handled by the complementary() function. The function starts with the current_rgb value, though it does not need to take it as an argument. First, the starting R, G, and B values are converted to HSV using the rgb_to_hsv(r,g,b) function. This function uses the colorsys rgb_to_hsv method. After finding the HSV for the complementary color, the color is converted back to RGB for easy use with the rest of the program functions. The HSV conversion is needed because the hue (H) value can be directly modified to find the new colors (finding a value 180, 30, and 120 degrees away on the hue circle). Subtracting the RGB values from 255 works to find the complementary colors, but it is more complex with the other color palette types.
3. Analogous will generate two colors that are 30 degrees away from the starting color in either direction on the color wheel.
- This is handled by the analogous() function. Adding and subtracting .083 and using modulo 1.0 finds the hue values that are 30 degrees away from the starting color.
4. Triadic will generate two colors that are 120 degrees away from the starting color in either direction on the color wheel.
- This is handled by the triadic() function. Adding and subtracting .33 and using modulo 1.0 finds the hue values that are 120 degrees away from the starting color.

- Select the desired palette type to generate a preview of the 2-3 colors. RGB values will also be generated and displayed below the color blocks.
  - Note that in Release v1.0, the RGB values are not able to be highlighted and copied.
<img width="1217" height="985" alt="Screenshot 2025-12-08 104516" src="https://github.com/user-attachments/assets/ebfef3b6-2435-4394-b895-11e549739b94" />

## Two options for templates to generate: An academic poster and presentation template. Both are in pptx format, and will save to the folder from which the user launched the program. Saving a new pptx file won't overwrite any existing files.

1. Poster:
<img width="1448" height="968" alt="Screenshot 2025-12-08 110739" src="https://github.com/user-attachments/assets/723dca46-10fe-4d2b-b571-7de14b9b823e" />
- The poster will be generated with the dimensions of 36x24 inches. If the conference requires a larger size, the user can modify the dimensions in PowerPoint by going to Design > Slide Size > Custom Slide Size and adjust the font sizes. Alternatively, some printers allow one to scale up the size of a document when printing. <b>Be sure to save as a PDF for best results!</b>
- This is handled by the poster() function. The function has sections for each part of the poster (title, introduction and space for text, methods and space for text, results and space for text, discussion and space for text, take home point, and references/contact information. Each box for the section title and text box is created as a retangle, with sizes hardcoded in inches. It uses the starting color and values generated by the palette radio buttons to set box colors. The poster will generate slightly differently depending on if there are two colors (complementary palette chosen) or three colors (analogous or triadic palettes chosen), which is explained in the if statement.
- The poster is saved as a .pptx file as "Poster + date and time (h, m, s) + .pptx".

<br>
2. Presentation:
<img width="1995" height="1007" alt="Screenshot 2025-12-08 104551" src="https://github.com/user-attachments/assets/47076cc5-c211-4a87-9e97-da14cfc2c3e1" />
- This is handled by the pres() function. As in the poster function, the colors are used to create a background (filled with starting color), text, and a line element. 
- The presentation is saved as a .pptx file as "Presentation + date and time (h, m, s) + .pptx".

## Additional Functions
1. setimage(): These lines of code are needed to reset the display box and newly display color swatches and images. It is placed into a function and called in the color selector and palette generator functions in order to limit redundancy.
2. force_bullet(paragraph, char): Creates bullet points in paragraph, used in the poster text sections and presentation.

## Known Issues
1. If the window is stretched length-wise, the RGB value for the middle color (self.values3) will move down. It is anchored to the window, while the other two labels are anchored to the grid. This was the only way I could think of to center the center value, because there are four columns in the UI and only three colors generated by the palette.
2. The RGB values cannot be highlighted or copied, because they are TkInter labels. There is a way to re-create them as text, but the process is involved.
3. Sometimes the generated colors are quite similar, or may not look professional enough. Always evaluate the poster and presentation templates yourself before use.

## Future Work
1. Implementing other color palettes, like tetradic and split complementary. This would likely just need a couple more palette functions and a slightly re-designed UI.
2. Making the RGB value labels into text that can be highlighted and copied.
3. Saving the RGB values for the palettes in a csv file, and stacking on when new ones are generated.
4. Presenting HSV, hex, etc. values alongside RGB values.
5. Expanding the functionality beyond posters and presentations--generating logos or even rooms with the colors from the palettes.
6. Allowing the user to choose a color from their entire screen, not just the uploaded image.
7. Creating a way to preview the poster and presentation templates in the display box before downloading.

(Python logo and color wheel image from Wikipedia.)

