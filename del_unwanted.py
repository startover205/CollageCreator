# This script is useful for cleaning the images dataset. It deletes all the
# corrupted images in a given folder leaving behind all the good images.
# Can be used to avoid the "OSError : cannot identify image file" while training
# the model.
#
#### WARNING: , any files other than image extensions accepted by PIL will be deleted.
#
# Usage : python del_unwanted.py $1
# where $1 = "dir_name" ( where images are stored )
# Requires absolute directory path
#
# Usage Example : python del_unwanted.py "~/downloaded"
#

import os
import sys
from PIL import Image
# import Pillow # let compiler install Pillow to use PIL


def remove_corrupted_image(dirpath):
    cnt=0
    # 遞迴開啟圖片檔，非圖片檔會被刪除！！
    for root, dirs, files in os.walk(dirpath):
        for file in files:
            filepath = os.path.join(root, file)
            try:
                # 刪除小於 1kb
                if os.path.getsize(filepath) < 1000:
                    raise OSError("File too small")
                # 刪除非圖片檔
                img = Image.open(filepath)
            except OSError:
                print("FILE: ", filepath, "is corrupted!")
                cnt += 1
                os.remove(filepath)
    print(f"Successfully Completed Operation! {cnt} Files Corrupted")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Error : Too Few Arguments \nUsage : python del_unwanted.py $1 \nwhere $1 = \"dir_name\" ( where images are stored ) \nRequires absolute directory path")
        sys.exit()
    dirname=sys.argv[1]
    cnt=0

    remove_corrupted_image(dirname)
    # remove_corrupted_image('/Users/mingtayang/PycharmProjects/CollageCreator/data')
