# Конвертирует все изображения gif, png или bmp в jpg по заданным параметрам

import os
import shutil
import sys

from PIL import Image


def convert_to_jpg(folder_name):
    output_folder = os.path.join(folder_name, os.path.basename(folder_name))
    os.makedirs(output_folder, exist_ok=True)
    for file in os.listdir(folder_name):
        if file.endswith(('.gif', '.png', '.bmp')):
            path = os.path.join(folder_name, file)
            img = Image.open(path)
            img.convert('RGB').save(os.path.join(output_folder, f'{os.path.splitext(file)[0]}.jpg'),
                                    quality=97, optimize=True, progressive=True, subsampling=0)
        elif file.endswith('.jpg'):
            shutil.copyfile(os.path.join(folder_name, file), os.path.join(output_folder, file))


if __name__ == '__main__':
    convert_to_jpg(sys.argv[1])
