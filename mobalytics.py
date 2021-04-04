# -*- coding: utf-8 -*-

# Basics
import copy
# Image processing
import imutils
import cv2
# Mouse and KeyBoard contollers
import win32api
import win32con
import win32com.client as comclt
# Miscellaneous
import pyautogui
import numpy as np

# https://pg.mobalytics.gg/

# For each, [xmin, ymin, xmax, ymax]
#   with each value being a ratio of the width or the height of the screen
# Those values have been computed with a 1920x1080 screen, on GoogleChrome
dim_to_click = [0.31223628691983124, 0.32375, 0.8094233473980309, 0.84625]
dim_key_contol = [0.19268635724331926, 0.3375, 0.30520393811533053, 0.57625]
dim_ship_avoider = [0.6744022503516175, 0.63125, 0.7988748241912799, 0.82875]

############################################################
####                                                    ####
####                 MOUSE INTERACTIONS                 ####
####                                                    ####
############################################################

# Coordinates of the mouse pointer
g_mouseX = g_mouseY = -1
# Coordinates of the mouse pointer when the left button is clicked
g_mouseX_click = g_mouseY_click = -1

# Click on the screen
def win32_click(x, y):
    # Move the cursor
    win32api.SetCursorPos((x, y))
    # Press the cursor
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    # Release the cursor
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

# OpenCV callback to get the cursor position on an OpenCV window
def get_position(event, x, y, flags, param):
    global g_mouseX, g_mouseY
    global g_mouseX_click, g_mouseY_click
    g_mouseX, g_mouseY = x, y
    if event == cv2.EVENT_LBUTTONDOWN:
        g_mouseX_click, g_mouseY_click = x, y

############################################################
####                                                    ####
#### USER INTERACTIONS TO GET THE REGIONS IN THE SCREEN ####
####                                                    ####
############################################################

def getDim():
    global g_mouseX, g_mouseY
    global g_mouseX_click, g_mouseY_click
    print("Getting the dimensions of the game'S sub-part on the screen...")
    print("Please follow the displayed instructions...\n")

    # Text hyper-paremeters
    thickness = 2
    fontScale = 1
    font = cv2.FONT_HERSHEY_SIMPLEX

    # First instructions
    text_ligne1 = '1) Launch the game'
    text_ligne2 = '2) Put this window in a corner'
    text_ligne3 = '3) Press a key when rdy'
    w_img, h_text = cv2.getTextSize(text_ligne2, font, fontScale, thickness)[0]

    # Display the first instructions
    img = np.zeros((3 * h_text + 30, w_img + 20))
    img = cv2.putText(img, text_ligne1, (5, 5 + h_text), font,
                      fontScale, 255, thickness, cv2.LINE_AA)
    img = cv2.putText(img, text_ligne2, (5, 15 + 2 * h_text), font,
                      fontScale, 255, thickness, cv2.LINE_AA)
    img = cv2.putText(img, text_ligne3, (5, 25 + 3 * h_text), font,
                      fontScale, 255, thickness, cv2.LINE_AA)
    cv2.imshow("Instructions", img)
    cv2.waitKey(0)
    cv2.destroyWindow("Instructions")

    # Do a screenshot
    screen = pyautogui.screenshot()
    # Convert the screenshot in an OpenCV understandable format
    screen = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2BGR)

    # Copy the screenshot
    img = copy.deepcopy(screen)

    # Next instructions
    instructions = [f'\'{instr}\' region: click on the {corner} corner...'
                    for instr in ('to-click', 'spell-control', 'ship-control')
                    for corner in ('top-left', 'bottom-right')]

    # Set a callback for the mouse use
    cv2.namedWindow('Screen')
    cv2.setMouseCallback('Screen', get_position)

    # Width of the instructions in the image
    w_text = 0
    # For all the instruction
    maxLenBeginLine = max(len(text[:text.find('region')-2]) for text in instructions)
    print(' ' * maxLenBeginLine + '    xmin\tymin\txmax\tymax')
    for i, text in enumerate(instructions):
        # Get the size of the next instruction to display
        w_texttmp, h_text = cv2.getTextSize(text, font, fontScale, thickness)[0]
        # Width of the instructions in the image
        w_text = max(w_text, w_texttmp)
        # Display the region where to put the instructions text
        img = cv2.rectangle(img,
                            (5, 5),
                            (8 + w_text, 16 + (i + 1) * h_text + 10 * i * (i>0)),
                            (0,0,0), -1)
        # Put the text of the current instructions and all its predecessors
        for j in range(i + 1):
            img = cv2.putText(img, instructions[j],
                              (5, 8 + (j + 1) * h_text + 10 * j * (j>0)),
                              font, fontScale, (255,255,255), thickness, cv2.LINE_AA)

        # Wait for a click
        g_mouseX_click = g_mouseY_click = -1
        while g_mouseX_click == -1 or g_mouseY_click == -1:
            # Resize the image
            resized_img = imutils.resize(img, height=800)
            h_resized, w_resized = resized_img.shape[:2]
            # Draw the cursor lines
            if g_mouseX != -1:
                cv2.line(resized_img, (g_mouseX, 0), (g_mouseX, h_resized), (0,0,0), 2)
            if g_mouseY != -1:
                cv2.line(resized_img, (0, g_mouseY), (w_resized, g_mouseY), (0,0,0), 2)
            # Show the screenshot with the cursor lines and the instructions
            cv2.imshow("Screen", resized_img)
            if 27 == cv2.waitKey(1):
                break

        # Print the clicks ration in the console
        if i%2 == 0:
            text_beginLine = text[:text.find('region')]
            text_beginLine = text_beginLine[1:-2]
            print('  ' + text_beginLine + ' ' * (maxLenBeginLine - len(text_beginLine)), end='')
            print(f'[ {round(g_mouseX_click / w_resized, 2)},\t{round(g_mouseY_click / h_resized, 2)},', end='\t')
        else:
            print(f'{round(g_mouseX_click / w_resized, 2)},\t{round(g_mouseY_click / h_resized, 2)}]')

    print("\nFeel free to update the sub-dims at the beginning of this script!")

#############################################################
####                                                     ####
####                  SPELLS CONTROLLER                  ####
####                                                     ####
#############################################################

# 1) Check the color of the bars
# 2) When the colors is a "go" signal, press the corresponding key
def spells_controllers(img, winshell, debug=False):
    # Get the image dimension
    h, w = img.shape[:2]
    # Convert in grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Check the 4 spells controllers
    for i in range(4):
        # If the color is a "go" signal
        if gray[(2*i+1)*h//8][w//3] > 220:
            # Press the corresponding key
            winshell.SendKeys('qwer'[i])
        if debug:
            # Draw the seperative horizontal lines for visual debug
            cv2.line(img, (0, (2*i+1)*h//8), (w, (2*i+1)*h//8), (0,0,255), 2)
    if debug:
        # Draw the seperative vertical line for visual debug
        cv2.line(img, (w//3, 0), (w//3, h), (0,255,0), 2)


############################################################
####                                                    ####
####                    SHIP AVOIDER                    ####
####                                                    ####
############################################################

# 1) Find the ship
# 2) Find a missile
# 3) If the ship is under the missile, make it move
def ship_avoider(img, pattern_ship, pattern_toavoid, winshell, threshold = 0.8):
    # Get the image dimension
    h, w = img.shape[:2]
    # Convert in grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Find the ship pattern
    res = cv2.matchTemplate(gray, pattern_ship, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)
    # Convert the loc into a list
    pts = list(zip(*loc[::-1]))

    # If no ship is found, return and do nothing
    if len(pts) < 1:
        return

    # Location of the ship: False is Right, True is Left
    sideShip = bool(pts[0][0] < w//2)

    # Find the missile pattern
    res = cv2.matchTemplate(gray, pattern_toavoid, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)
    # Convert the loc into a list
    pts = list(zip(*loc[::-1]))

    # If no missile is found, return and do nothing
    if len(pts) < 1:
        return

    # Location of the missile: False is Right, True is Left
    sideMissile = bool(pts[0][0] < w//2)

    # If the ship and the missile have the same location
    if sideShip == sideMissile:
        # Make the ship move
        winshell.SendKeys('df'[sideShip])


############################################################
####                                                    ####
####                    AUTO-CLICKER                    ####
####                                                    ####
############################################################

# 1) Find all the clickable patterns
# 2) Click on them
def auto_clicker(img, pattern, offsetX, offsetY, threshold=0.8, debug=False):
    # Get the pattern dimension
    w, h = pattern.shape[::-1]
    # Convert in grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # Find the pattern
    res = cv2.matchTemplate(gray, pattern, cv2.TM_CCOEFF_NORMED)
    loc = np.where(res >= threshold)
    # Convert the loc into a list
    pts = list(zip(*loc[::-1]))

    # For all the detected patterns
    for pt in pts:
        # Click on the center of the detected pattern
        win32_click(offsetX + pt[0] + w //2 , offsetY + pt[1] + h//2)
        if debug:
            # Draw the detected pattern for visual debug
            cv2.rectangle(img, pt, (pt[0] + w, pt[1] + h), (0,0,255), 2)

    return len(pts) > 0


############################################################
####                                                    ####
####                        MAIN                        ####
####                                                    ####
############################################################

if __name__ == '__main__':
    # Do again the user intersection to get the windows dimensions
    doGetDim = True
    if doGetDim or dim_to_click == -1 or dim_key_contol == -1 or dim_ship_avoider == -1:
        getDim()
    else:
        # Read in grayscale the pattern images used in the different algorithms
        print("Loading the patterns images...")
        pattern_ship = cv2.imread("data/patternShip.png", 0)  # Pattern of the ship
        pattern_ship_toavoid = cv2.imread("data/patternShipToAvoid.png", 0)  # Pattern of a missile
        pattern_toClick = cv2.imread("data/patternToClick.png", 0)  # Pattern of a clickable area
        print("Loading the patterns images...  DONE")

        print("\nGetting a the screen dimensions...")
        # Do a first screenshot to get the dimension
        image = pyautogui.screenshot()
        # Convert the screenshot in an OpenCV understandable format
        image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        # Get the screen dimension
        wxh = image.shape[:2][::-1]
        print(f"    Screen dimensions -> {wxh[0]}x{wxh[1]}")
        print("Getting a the screen dimensions...  DONE")

        print("\nPreparing the sub-dimensions...")
        # Update the sub-dimensions values
        print("            \txmin\tymin\txmax\tymax")
        for dim in (dim_to_click, dim_key_contol, dim_ship_avoider):
            for i in range(len(dim)):
                dim[i] = round(dim[i] * wxh[i%2])
            print(f"  sub-dim ->\t{dim[0]}\t{dim[1]}\t{dim[2]}\t{dim[3]}")
        print("Preparing the sub-dimensions...  DONE")

        # WinShell for keyboard output
        winshell = comclt.Dispatch("WScript.Shell")
        winshell.AppActivate("Notepad")

        print("\nApp ready, you can launch the game...")
        print("Press CTRL+C to stop this app...")
        # Try to catch the "CTRL+C" keyboard event
        try:
            onGame = False  # Whether or not the game is ongoing
            countNoClick = 0  # Number of successive screenshots in which no auto-click is performed
            while True:
                # Do a screenshot
                image = pyautogui.screenshot()
                # Convert the screenshot in an OpenCV understandable format
                image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                #cv2.imshow("image", imutils.resize(image, height=800))

                # Extract region where to apply the auto-clicker
                img_to_click = image[dim_to_click[1]:dim_to_click[3],dim_to_click[0]:dim_to_click[2]]
                # Apply the auto-clicker
                if auto_clicker(img_to_click, pattern_toClick, dim_to_click[0], dim_to_click[1]):
                    if not onGame:
                        print("Game is ON!")
                    onGame = True
                    countNoClick = 0
                else:
                    countNoClick += 1
                #cv2.imshow("to_click", img_to_click)

                if countNoClick > 15:
                    if onGame:
                        print("Game is OFF!")
                    onGame = False

                if onGame:
                    # Extract region where to apply the spell contollers
                    img_spells_control = image[dim_key_contol[1]:dim_key_contol[3],dim_key_contol[0]:dim_key_contol[2]]
                    # Apply the spell contollers
                    spells_controllers(img_spells_control, winshell)
                    #cv2.imshow("spells_control", img_spells_control)

                    # Extract region where to apply the ship avoider
                    img_ship_avoider = image[dim_ship_avoider[1]:dim_ship_avoider[3],dim_ship_avoider[0]:dim_ship_avoider[2]]
                    # Apply the ship avoider
                    ship_avoider(img_ship_avoider, pattern_ship, pattern_ship_toavoid, winshell)
                    #cv2.imshow("ship_avoider", img_ship_avoider)

                # If an OpenCV window is open, close the app by pressing 'ESC'
                if 27 == cv2.waitKey(1):
                    break
        # Catch the "CTRL+C" keyboard event
        except KeyboardInterrupt:
            print('Ctrl-C pressed, killing the app...')

    cv2.destroyAllWindows()
