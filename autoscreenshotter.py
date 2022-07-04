import pyautogui
import scipy
from PIL import ImageChops, Image
import time

from io import BytesIO
import win32clipboard

import sys
import cv2
from scipy import stats
from scipy.linalg import norm

import numpy as np

SS_size_x, SS_size_y = (900, 510) # Default is 900, 510
SS_loc_offset_x, SS_loc_offset_y = (50, 280) # Default is 50, 280

pause_time = 3.5 # seconds

zoom_pic_file_loc = 'pics\zoomicon.png'
Word_save_icon_file_loc = 'pics\saveicon.png'

def send_to_clipboard(PIL_image):
    output = BytesIO()
    PIL_image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def screenshot_and_save_to_word(screenshot):
    old_cursor_position = pyautogui.position()
    send_to_clipboard(screenshot)
    print(f"Looking for image on screen, returning output: ")
    loc = pyautogui.locateOnScreen(Word_save_icon_file_loc, grayscale=False, confidence=.9) # returns (left, top, width, height) of first place it is found
    print(loc)
    pyautogui.moveTo(loc)
    # pyautogui.moveRel(0, 0, duration=0.1)
    pyautogui.click(clicks=2, interval=0, button='left') #
    pyautogui.moveTo(old_cursor_position)
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl', 'v')



## SOURCE: https://stackoverflow.com/questions/189943/how-can-i-quantify-difference-between-two-images
def to_grayscale(arr):
    "If arr is a color image (3D array), convert it to grayscale (2D array)."
    print(f"arr: {arr}")
    print(f"arr: {arr}")
    if len(arr.shape) == 3:
        return np.average(arr, -1)  # average over the last axis (color channels)
    else:
        return arr

def normalize(arr):
    # print(f"arr in normalize: {arr}")
    # print(f"arr.max: {arr.max()}")
    # print(f"arr.min: {arr.min()}")
    rng = arr.max()-arr.min()
    # print(f"rng: {rng}")
    amin = arr.min()
    # print(f"amin: {amin}")
    # print(f"arr-amin: {arr-amin}")
    # print(f"arr-amin * 255: {(arr - amin)*4}")
    # print(f"FINAL RESULT: {(arr-amin)*255/rng}") # When multipled by 255, any value above 255 gets wrapped around back to 0 (zero itself is takes up a value of 1). See https://stackoverflow.com/questions/40982605/why-is-the-max-element-value-in-a-numpy-array-255
    return (arr-amin)*255/rng


def compare_images(img_1, img_2): ## img_1 and img_2 are PIL images
    # Convert PIL to arrays using numpy first
    img_1_arr = np.array(img_1)
    img_2_arr = np.array(img_2)
    # print(f"img_1_arr: {img_1_arr}")
    # print(f"img_2_arr: {img_2_arr}")
    # normalize to compensate for exposure difference, this may be unnecessary
    # consider disabling it
    img_1_normalized = normalize(img_1_arr)
    img_2_normalized = normalize(img_2_arr)
    # print(f"img_1_normalized: {img_1_normalized}")
    # print(f"img_2_normalized: {img_2_normalized}")
    # calculate the difference and its norms
    diff_normalized = img_1_normalized - img_2_normalized  # elementwise for scipy arrays
    # print(f"diff_normalized: {diff_normalized}")
    m_norm = sum(abs(diff_normalized))  # Manhattan norm
    # print(f"Manhattan norm without sum: {abs(diff_normalized)}")
    print(f"Manhattan norm: {m_norm}")
    z_norm = norm(diff_normalized.ravel(), 0)  # Zero norm. .ravel() combines arrays into a single array. norm() calculates the Frobenius norm (aka Euclidean norm) of an array, and it will be the same even if the single array was reshaped into multiple arrays.
    print(f"Zero norm: {z_norm}")
    return (m_norm, z_norm)

# def compare_images(img_1, img_2):
    # # --- take the absolute difference of the images ---
    # res = cv2.absdiff(img_1, img_2)
    #
    # # --- convert the result to integer type ---
    # res = res.astype(np.uint8)
    #
    # # --- find percentage difference based on number of pixels that are not zero ---
    # percentage_diff = (np.count_nonzero(res) * 100) / res.size
    # print(f"PERCENTAGE DIFFERENCE BTWN IMAGES: {percentage_diff}")


if __name__ == "__main__":
    print(f"Current screen size is {pyautogui.size()}\n") # current screen resolution width and height
    # pyautogui.PAUSE = 2.5 # Set up a 2.5 second pause after each PyAutoGUI call
    pyautogui.FAILSAFE = True # When fail-safe mode is True, moving the mouse to the upper-left will raise a pyautogui.FailSafeException that can abort your program:
      # current mouse x and y
    screenshot_count = 0
    while screenshot_count == 0:
        print(f"Looking for zoom image on screen: ")
        loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False, confidence=.9) # returns (left, top, width, height) of first place it is found
        # print(loc2)
        # pyautogui.moveTo(loc2) #, duration=0.1)
        if loc is None:
            print(f"Zoom image not found. Pausing before trying again\n")
            time.sleep(pause_time)
            continue
        im = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))
        screenshot_and_save_to_word(im)
        screenshot_count += 1
        print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")

    while screenshot_count > 0:
        # pyautogui.PAUSE = pause_time
        time.sleep(pause_time)
        print("====Checking if new screenshot is different:====")
        print(f"Looking for zoom image on screen: ")
        loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False,
                                        confidence=.9)  # returns (left, top, width, height) of first place it is found
        # print(loc2)
        # pyautogui.moveTo(loc2) #, duration=0.1)
        if loc is None:
            print(f"Zoom image not found. Pausing before trying again\n")
            continue
        im_next = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))
        m_norm, z_norm = compare_images(im, im_next)
        if z_norm < 5000: # Next screenshot is similar to previous screenshot
            print("...New screenshot is identical to old screenshot. Pausing before checking again...\n")
            continue
        else:
            im = im_next
            screenshot_and_save_to_word(im)
            screenshot_count += 1
            print(f"-----New screenshot is different to old screenshot.-----")
            print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")

            continue
        break
        # diff = ImageChops.difference(im, im_next)
        # bbox = diff.getbbox()
        # if bbox is not None: # exact comparison
        #     im = im_next
        #     screenshot_and_save_to_word(im)
        #     print("-----New screenshot is different to old screenshot. Pasting in Word doc-----")
        #     continue
        # else:
        #     print("...New screenshot is identical to old screenshot. Pausing before checking again...")


    # print("bounding box of non-zero difference: %s" % (bbox,))
    # # superimpose the inverted image and the difference
    # ImageChops.screen(ImageChops.invert(im.crop(bbox)), diff.crop(bbox)).show()
    # input("Press Enter to exit")