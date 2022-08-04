import copy

import pyautogui
import keyboard
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


HOTKEY_MODE = False
hotkey_to_screenshot = 'ctrl+l'

## Options to modify image arrays before comparing
black_threshold = 50 ## Default = 50
percentage_black_threshold = 0.987

SS_size_x, SS_size_y = (950, 580) # Default 740, 550
SS_loc_offset_x, SS_loc_offset_y = (50, 130) # Default is 50, 280 #120, 220

pause_time = 4.5 # seconds

resize_img_smaller = False
resize_to = (100, 56)
convert_img_to_grayscale = False
normalise_img = False

## USELESS VARIABLES - IGNORE
# z_norm_diff_cutoff = 1900
fail_safe_boolean = False # When fail-safe mode is True, moving the mouse to the upper-left will raise a pyautogui.FailSafeException that can abort your program:

zoom_pic_file_loc = 'pics\zoomicon.png'
Word_save_icon_file_loc = 'pics\saveicon.png'

def send_to_clipboard(PIL_img):
    output = BytesIO()
    PIL_img.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def screenshot_and_save_to_word(PIL_img):
    old_cursor_position = pyautogui.position()
    send_to_clipboard(PIL_img)
    loc = pyautogui.locateOnScreen(Word_save_icon_file_loc, grayscale=False, confidence=.9) # returns (left, top, width, height) of first place it is found
    print(f"Looking for Word Doc image on screen, returning output: {loc}")
    pyautogui.moveTo(loc)
    # pyautogui.moveRel(0, 0, duration=0.1)
    pyautogui.click(clicks=1, interval=0, button='left')
    pyautogui.moveTo(old_cursor_position)
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl', 'v')



## SOURCE: https://stackoverflow.com/questions/189943/how-can-i-quantify-difference-between-two-images
def to_grayscale(arr):
    "If arr is a color image (3D array), convert it to grayscale (2D array)."
    # print(f"arr: {arr}")
    # print(f"arr: {arr}")
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


def compare_images_using_norms(PIL_img_1, PIL_img_2): ## img_1 and img_2 are PIL images
    # Convert PIL to arrays using numpy first
    img_1_arr = np.array(PIL_img_1)
    img_2_arr = np.array(PIL_img_2)
    # print(f"img_1_arr: {img_1_arr}")
    # print(f"img_2_arr: {img_2_arr}")

    # ## grayscale
    if convert_img_to_grayscale is True:
        print(f"Converting images to grayscale...")
        img_1_arr = to_grayscale(img_1_arr)
        img_2_arr = to_grayscale(img_2_arr)
        # print(f"img_1_arr_grayscale: {img_1_arr}")
        # print(f"img_2_arr_grayscale: {img_2_arr}")

    ## normalize to compensate for exposure difference, this may do more wrong than good. For example, a single bright pixel on a dark background will make the normalized image very different
    ## consider disabling it
    if normalise_img is True:
        print(f"Normalizing images...")
        img_1_arr = normalize(img_1_arr)
        img_2_arr = normalize(img_2_arr)
        # print(f"img_1_arr_normalized: {img_1_arr}")
        # print(f"img_2_arr_normalized: {img_2_arr}")

    ## calculate the difference and its norms
    diff = img_1_arr - img_2_arr  # elementwise for scipy arrays
    # print(f"diff_in_arrays: {diff}")
    m_norm = sum(abs(diff))  # Calculate Manhattan norm of diff
    # print(f"Manhattan norm without sum: {abs(diff_normalized)}")
    # print(f"Manhattan norm: {m_norm}")
    z_norm = norm(diff.ravel(), 0)  # Calculate Zero norm of diff; .ravel() combines arrays into a single array. norm() calculates the Frobenius norm (aka Euclidean norm) of an array, and it will be the same even if the single array was reshaped into multiple arrays.
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


def resize_images(PIL_img): ## accepts non-array images
    img_array = np.array(PIL_img)
    resized_image = cv2.resize(img_array, resize_to, interpolation=cv2.INTER_LINEAR)
    return resized_image

def return_diff_of_images_with_ImageChops(PIL_img_1, PIL_img_2): ## returns a PIL image showing the difference in pixels (if mostly similar, returns mostly black)
    if convert_img_to_grayscale is True:
        print(f"Converting images to grayscale...")
        PIL_img_1 = to_grayscale(PIL_img_1)
        PIL_img_2 = to_grayscale(PIL_img_2)

    ## normalize to compensate for exposure difference, this may do more wrong than good. For example, a single bright pixel on a dark background will make the normalized image very different
    ## consider disabling it
    if normalise_img is True:
        print(f"Normalizing images...")
        PIL_img_1 = normalize(PIL_img_1)
        PIL_img_2 = normalize(PIL_img_2)

    diff = ImageChops.difference(PIL_img_1, PIL_img_2)
    print(f"diff: {diff}")

    ## Bounding box method en-boxes all non-zero pixels, and returns a Box PIL object that contains the 3D dimensions
    ## However, this method is too sensitive; even when images are mostly similar, the Box dimensions could still be very big!
    # bbox = diff.getbbox() ## bounding box
    # print(f"bbox: {bbox}")
    # if bbox is None:  # Next screenshot is similar to previous screenshot
    #     print("...New screenshot is identical to old screenshot. Pausing before checking again...\n")
    #     continue
    return diff


def calculate_percentage_black(PIL_img, black_threshold):
    pixels = PIL_img.getdata()  # get the pixels as a flattened sequence
    nblack = 0
    for pixel in pixels:
        if sum(pixel) < black_threshold:
            nblack += 1
    n = len(pixels)
    percentage_black = (nblack / float(n))
    print(f"nblack: {nblack}; float(n): {float(n)}; %_black: {percentage_black}")
    return percentage_black


if __name__ == "__main__":
    print(f"Current screen size is {pyautogui.size()}\n") # current screen resolution width and height
    # pyautogui.PAUSE = 2.5 # Set up a 2.5 second pause after each PyAutoGUI call
    pyautogui.FAILSAFE = fail_safe_boolean # When fail-safe mode is True, moving the mouse to the upper-left will raise a pyautogui.FailSafeException that can abort your program:
      # current mouse x and y
    screenshot_count = 0
    if HOTKEY_MODE is True:
        while True:
            time.sleep(0.1)
            print(f"HOTKEY_MODE is True: press {hotkey_to_screenshot} to screenshot!")
            try:  # used try so that if user pressed other than the given key error will not be shown
                if keyboard.is_pressed(hotkey_to_screenshot):  # if key 'q' is pressed
                    print('You pressed a Key!')
                    loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False, confidence=.9)  # returns (left, top, width, height) of first place it is found
                    if loc is None:
                        print(f"Zoom image not found. \n")
                        continue
                    im = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))
                    screenshot_and_save_to_word(im)
                    screenshot_count += 1
                    print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")
                    continue  # finishing the loop
            except:
                continue  # if user pressed a key other than the given key the loop will continue/break
    while screenshot_count == 0:
        loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False, confidence=.9) # returns (left, top, width, height) of first place it is found
        print(f"Looking for zoom image on screen: {loc}")
        if loc is None:
            print(f"Zoom image not found. Pausing before trying again\n")
            time.sleep(pause_time)
            continue
        im = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))
        screenshot_and_save_to_word(im)
        screenshot_count += 1
        print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")

    # while screenshot_count > 0:
    #     # pyautogui.PAUSE = pause_time
    #     time.sleep(pause_time)
    #     print("====Checking if new screenshot is different:====")
    #     print(f"Looking for zoom image on screen: ")
    #     loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False, confidence=.9)  # returns (left, top, width, height) of first place it is found
    #     if loc is None:
    #         print(f"Zoom image not found. Pausing before trying again\n")
    #         continue
    #     im_next = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))
    #
    #     if resize_img_smaller is False:
    #         m_norm, z_norm = compare_images_using_norms(im, im_next)
    #     else:
    #         im_downsized = resize_images(im)
    #         im_next_downsized = resize_images(im_next)
    #         m_norm, z_norm = compare_images_using_norms(im_downsized, im_next_downsized)
    #
    #     if z_norm < z_norm_diff_cutoff: # Next screenshot is similar to previous screenshot
    #         print("...New screenshot is identical to old screenshot. Pausing before checking again...\n")
    #         continue
    #     else:
    #         im = copy.deepcopy(im_next)
    #         screenshot_and_save_to_word(im)
    #         screenshot_count += 1
    #         print(f"-----New screenshot != previous screenshot.-----")
    #         print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")
    #         continue


    while screenshot_count > 0:
        # pyautogui.PAUSE = pause_time
        time.sleep(pause_time)
        print("====Checking if new screenshot is different:====")
        loc = pyautogui.locateOnScreen(zoom_pic_file_loc, grayscale=False,confidence=.9)  # returns (left, top, width, height) of first place it is found
        print(f"Looking for zoom image on screen: {loc}")
        if loc is None:
            print(f"Zoom image not found. Pausing before trying again\n")
            continue
        im_next = pyautogui.screenshot(region=(loc[0] + SS_loc_offset_x, loc[1] + SS_loc_offset_y, SS_size_x, SS_size_y))

        if resize_img_smaller is True:
            im_next_downsized = resize_images(im_next)
            im_downsized = resize_images(im)
            diff = return_diff_of_images_with_ImageChops(im_downsized, im_next_downsized)
        else:
            diff = return_diff_of_images_with_ImageChops(im, im_next)
        percentage_black = calculate_percentage_black(diff, black_threshold)
        if percentage_black < percentage_black_threshold:
            im = copy.deepcopy(im_next)
            screenshot_and_save_to_word(im_next)
            screenshot_count += 1
            print(f"-----%_black: {percentage_black}; threshold reached! New screenshot != previous screenshot.-----")
            print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")
            continue
        else:
            print("...New screenshot is identical to old screenshot. Pausing before checking again...\n")
            continue


            # if resize_img_smaller is True:
            #     im_downsized = resize_images(im)
            #
            # screenshot_count += 1
            # print(f"-----New screenshot is different to old screenshot.-----")
            # print(f"-----Screenshot No. {screenshot_count}  is saved-----\n")





        # diff = ImageChops.difference(im, im_next)
        # bbox = diff.getbbox()
        # if bbox is not None: # exact comparison
        #     im = copy.deepcopy(im_next)
        #     screenshot_and_save_to_word(im)
        #     print("-----New screenshot is different to old screenshot. Pasting in Word doc-----")
        #     continue
        # else:
        #     print("...New screenshot is identical to old screenshot. Pausing before checking again...")


    # print("bounding box of non-zero difference: %s" % (bbox,))
    # # superimpose the inverted image and the difference
    # ImageChops.screen(ImageChops.invert(im.crop(bbox)), diff.crop(bbox)).show()
    # input("Press Enter to exit")