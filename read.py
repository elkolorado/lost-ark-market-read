import pytesseract
import pandas as pd
from PIL import Image
import re
import time
import os
import win32clipboard
from io import BytesIO

pytesseract.pytesseract.tesseract_cmd = r'G:\tesseract\tesseract.exe'


def ocr_core(image, type='string'):
    """
    This function will handle the core OCR processing of images.
    """

    config = ('-l eng --oem 1 --psm 6')  # English language, OEM Engine mode 1, Page Segmentation Mode 6
    text = pytesseract.image_to_string(image, config=config)  # We'll use Pillow's Image class to open the image and pytesseract to detect the string in the image

    # get only first line of the text
    # text = text.split("\n")[0]

    #remove \n from the text, and leave just letters
    text = re.sub(r'\n', '', text)
    if(type == 'string'):
        text = re.sub(r'[^A-Za-z ]', '', text)
        #trim the text
        text = text.strip()
    if(type == 'number'):
        text = re.sub(r'[^0-9]', '', text)
    if(type == 'number_decimal'):
        #rerrun the OCR with different config, to get the decimal number
        config = ('-l eng --oem 1 --psm 6 -c tessedit_char_whitelist=0123456789.')
        # Increase DPI to improve OCR accuracy for small dots
        image = image.resize((image.width * 4, image.height * 4), Image.Resampling.LANCZOS)
        # image = image.convert("L")  # Convert to grayscale to improve OCR accuracy
        config = ('-l eng --oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.')  # Use OEM 3 for better accuracy
        text = pytesseract.image_to_string(image, config=config)
        text = re.sub(r'[^0-9]', '', text)
        text = text.strip()


        
    return text


#crop the table in the image X: 475, Y: 332, height: 385, width: 991
table_x = 920
table_y = 262
table_height = 385+43
table_width = 991
row_height = 43

name_width = 444
name_crop_width = 200


price_width = 128

lowest_price_width = name_width+price_width*2
lowest_price_crop_width = name_width+price_width+price_width*2


def tableCrop(img_path):    
    img = Image.open(img_path)
    img = img.crop((table_x, table_y, table_x+table_width, table_y+table_height))
    # img.save("table.png")
    return img



# Function to iterate through the table by row height and call a crop function for each row
def processTableRows(cropFunction, img):
    results = []  # List to store results from the crop function
    for i in range(0, table_height, row_height):
        names = cropFunction(i, img)
        results.append(names)
    return results


#crop the rows in the table, start from X: 0, Y: 0, height: 43, width: 444, and then add 43 to Y
def nameCrop(i, img):

    img1 = img.crop((0, i, name_width, i+row_height))
    # img1.save(f"row{i}.png")

    # Crop the name area from row{i}.png
    # img2 = Image.open(f"row{i}.png")
    img2 = img1.crop((row_height-1, 0, name_crop_width, row_height-22))
    # img2.save(f"name{i}.png")
    ocrd = ocr_core(img2)

    if len(re.sub(r'\W+', '', ocrd)) <= 2:
        # Expand crop area if OCR result is too short
        img2 = img2.crop((0, 0, name_crop_width, row_height-22+10))
        # img2.save(f"name{i}.png")
        ocrd = ocr_core(img2)

    return ocrd 


def lowestPrice(i, img):
    img1 = img.crop((lowest_price_width, i, lowest_price_crop_width, i+row_height))
    # img1.save(f"row{i}.png")

    ocrd = ocr_core(img1, type='number_decimal')

    return ocrd


def recentPrice(i, img):
    img1 = img.crop((lowest_price_width-price_width, i, lowest_price_crop_width-price_width, i+row_height))
    # img1.save(f"row{i}.png")

    ocrd = ocr_core(img1, type='number_decimal')

    return ocrd



def processScreen(image_path):
    img = tableCrop(image_path)  # Call the function to crop the table
    names = processTableRows(nameCrop, img)  # Call the function to process the table rows
    print(names)  # Print the results

    recent_prices = processTableRows(recentPrice, img)  # Call the function to process the table rows
    print(recent_prices)  # Print the results

    prices = processTableRows(lowestPrice, img)  # Call the function to process the table rows
    print(prices)  # Print the results

    # Load existing data from the CSV file if it exists


    try:
        df_existing = pd.read_csv("output.csv")
    except FileNotFoundError:
        df_existing = pd.DataFrame(columns=['Name', 'Recent Price', 'Lowest Price'])

    # Create a DataFrame from the results
    df_new = pd.DataFrame(list(zip(names, recent_prices, prices)), columns=['Name', 'Recent Price', 'Lowest Price'])

    # Ensure numeric values and fill missing values with 0
    df_new[['Recent Price', 'Lowest Price']] = df_new[['Recent Price', 'Lowest Price']].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    # Update the existing DataFrame with new data
    df_updated = pd.concat([df_existing.set_index('Name'), df_new.set_index('Name')], axis=0)
    df_updated = df_updated.groupby('Name').last().reset_index()

    # Ensure numeric values again for the final DataFrame
    df_updated[['Recent Price', 'Lowest Price']] = df_updated[['Recent Price', 'Lowest Price']].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    # Save the updated DataFrame to the CSV file
    df_updated.to_csv("output.csv", index=False)





# processScreen("input.jpg") 
# processScreen("input2.jpg")

from g import update_spreadsheet_from_csv  # Import the function from the g.py file
 # Call the function to update the Google Sheets spreadsheet


#instead we will listen to screenshots and process them
def listen_to_print_screen():
    """
    Listen for the Print Screen hotkey and process the image from the clipboard.
    """
    print("Listening for Print Screen hotkey...")

    processed_images = set()  # Keep track of already processed images

    while True:
        # Check if there's an image in the clipboard
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                # Convert DIB data to BMP format for PIL compatibility
                bmp_header = b'BM' + (len(clipboard_data) + 14).to_bytes(4, 'little') + b'\x00\x00\x00\x00\x36\x00\x00\x00'
                bmp_data = bmp_header + clipboard_data
                image = Image.open(BytesIO(bmp_data))
                temp_image_path = "temp_image.png"
                image.save(temp_image_path)  # Save the image to a temporary file


                # Generate a unique identifier for the image
                image_id = hash(image.tobytes())

                if image_id not in processed_images:
                    processed_images.add(image_id)  # Mark the image as processed
                    print("Processing new image from clipboard...")
                    processScreen(temp_image_path)  # Pass the file path to processScreen
                    update_spreadsheet_from_csv()  # Update the Google Sheets spreadsheet

                os.remove(temp_image_path)  # Remove the temporary file after processing

        except Exception as e:
            print(f"Error processing clipboard image: {e}")
        finally:
            win32clipboard.CloseClipboard()

        time.sleep(1)  # Check the clipboard every second


# Start listening to the screenshots directory
if __name__ == "__main__":
    listen_to_print_screen()




# Read the data from the worksheet
