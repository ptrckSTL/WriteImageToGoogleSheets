import gspread
import gspread_formatting
import string
from PIL import Image
from gspread_formatting import *


def decode_image_data():
    im = Image.open('image.png', 'r')
    pixels = get_image_pixels(list(im.getdata()))
    index = 0
    columns, rows = im.size
    result = []
    for row in range(rows):
        for column in range(columns):
            result.append([str(convert_to_a1_notation(column + 1)) + str(row + 1), pixels[index]])
            index += 1
    write_to_sheet(rows, columns, result)


def get_image_pixels(pixel_values):
    data = []
    for pixel in pixel_values:
        r, g, b, a = pixel
        data.append(CellFormat(backgroundColor=Color.fromHex(hex_triplet([r, g, b]))))
    return data


def hex_triplet(colortuple):
    return '#' + ''.join(f'{i:02X}' for i in colortuple)


def convert_to_a1_notation(num):
    title = ''
    alist = string.ascii_uppercase
    while num:
        mod = (num - 1) % 26
        num = int((num - mod) / 26)
        title += alist[mod]
    return title[::-1]


def write_to_sheet(rows, columns, pixel_coordinates):
    gc = gspread.oauth()
    sh = gc.open_by_key("1f-tLLV-oQz4ykqllMyd6bEErmpxpYhuREGGPbKeNKZE")
    worksheet = sh.worksheet("Sheet1")
    worksheet.resize(rows, columns)
    formatter = gspread_formatting

    set_column_width(worksheet, "A:" + convert_to_a1_notation(columns), 2)
    set_row_height(worksheet, "1:" + str(rows), 3)
    formatter.format_cell_ranges(worksheet, pixel_coordinates)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    decode_image_data()

