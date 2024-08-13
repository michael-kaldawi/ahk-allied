# Automation Steps

## Badge Checkout

### Manual Process

  1. Open Camera desktop app.
  2. Take a close up photo of the badge.
  3. Open the images by clicking on the most recently taken photo.
  4. Press ctrl+shift+a to open the Lark .screenshot utility, and click "Image to Text".
  5. Extract the badge ID from the image visible in the camera app.
  6. Paste the ID into the spreadsheet (ctrl+v).

### AutoHotKey Process

  The autohotkey script, converted to an executable, should be placed in the startup folder to run at system startup. The location of the startup folder can be reached by pressing Windows Key + R, and then entering shell:startup. 

  1. Open the Seattle badges distribution log window, and select the appropriate sheet (Temporary or Short-Term, depending on the type of badge checkout).
  2. Open/activate the Camera app window. The camera should be set in QR scanning mode.
  3. Hold the badge up to the camera. The camera will scan the badge, copy the badge number, and paste it into the spreadsheet along with the current date.
