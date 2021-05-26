import cv2 as cv
import numpy as np
import os
import random
import copy
from fpdf import FPDF
import download_with_urllib
from multiprocessing import Pool


# 畫上圖片
def draw_on(img, img2, x, y):
    img[y:y + img2.shape[0], x: x + img2.shape[1]] = img2


# scale to fit
def scale_to_fit(img, w, h):
    img_w = img.shape[1]
    img_h = img.shape[0]

    ratio_w = w / img_w

    result_h = img_h * ratio_w
    if result_h > h:
        ratio_h = h / img_h
        result_w = img_w * ratio_h
        return cv.resize(img, (int(result_w), int(h)))
    else:
        return cv.resize(img, (int(w), int(result_h)))


if __name__ == '__main__':
    # 定義項目
    animals = ['cow', 'pig', 'cat', 'fish', 'chicken', 'dog', 'duck','horse']
    clothes = ['socks', '短袖', 'pants', 'shoes', 'bed']
    body_parts = ['ear', 'nose', 'mouth', 'eyes', 'feet', 'hand']
    # people = ['grandpa', 'grandpa2', 'uncle', 'uncle2', 'sister', 'dad', 'great_grandma', 'grandma', 'mother']
    furniture = ['desk', 'chair']
    vehicles = ['car', 'motorcycle']
    nature = ['rock', 'flower', 'grass', 'tree']
    fruits = ['banana', 'orange', 'apple', 'guava', '藍莓', 'grape']
    vegetables = ['carrot']
    cartoons = ['巧虎']
    treats = ['snacks', '饅頭', '點心麵']
    life_stuff = ['soap', 'diaper', '牙刷']
    instruments = ['drum']

    # XXX colors = ['red', 'green']

    everything = [animals, clothes, body_parts, furniture, vehicles, nature, fruits, vegetables, cartoons, treats, life_stuff, instruments]


    # 定義下載項目
    target = ['dog']
    similars = animals

    main = [item for sub in everything for item in sub]
    supplemented = ['cartoon']

    # 下載圖片
    path = './data/'
    # download_with_urllib.download(path, main, supplemented)

    # 讀取圖片
    # 卡通
    # 1~17: 1 張狗, 2 張其他
    # 18~34: 1 張狗, 1 張動物, 1 張其他
    # 34~50: 1 張狗, 2 張動物

    # 分三區：卡通狗、非狗卡通動物、其他
    img_target = []
    img_similars = []
    img_unsimilars = []

    for dirname, subdirnames, filenames in os.walk(path):
        for filename in filenames:
            if filename.endswith('jpg') or filename.endswith('png'):
                fn = os.path.join(dirname, filename)
                if os.path.basename(dirname) in target:
                    img_target.append(fn)
                elif os.path.basename(dirname) in similars:
                    img_similars.append(fn)
                else:
                    img_unsimilars.append(fn)

    print(similars)

    random.shuffle(img_target)
    random.shuffle(img_similars)
    random.shuffle(img_unsimilars)
    print(len(img_target), len(img_similars), len(img_unsimilars))

    # 分類成列
    # 1~17: 1 張狗, 2 張其他
    # 18~34: 1 張狗, 1 張動物, 1 張其他
    # 34~50: 1 張狗, 2 張動物

    rows = []

    # 補足缺少的圖片數量
    img_target_copy = copy.deepcopy(img_target)

    for _ in range(17):
        row = [img_target.pop(), img_unsimilars.pop(), img_unsimilars.pop(), img_unsimilars.pop(), img_unsimilars.pop()]
        random.shuffle(row)
        rows.append(row)
        if len(img_target) < 1:
            img_target.extend(img_target_copy)

    for row in rows:
        print(row)
    print('--------------')

    for _ in range(17, 34):
        row = [img_target.pop(), img_similars.pop(), img_unsimilars.pop(), img_unsimilars.pop(), img_unsimilars.pop()]
        random.shuffle(row)
        rows.append(row)
        if len(img_target) < 1:
            img_target.extend(img_target_copy)

    for row in rows:
        print(row)
    print('--------------')

    for _ in range(34, 50):
        row = [img_target.pop(), img_similars.pop(), img_similars.pop(), img_unsimilars.pop(), img_unsimilars.pop()]
        random.shuffle(row)
        rows.append(row)
        if len(img_target) < 1:
            img_target.extend(img_target_copy)

    for row in rows:
        print(row)
    print('--------------')

    assert len(rows) == 50

    # 排列圖列拼貼

    # 建立 A4 圖片
    # A4 2479×3508
    total_width = 3508
    total_height = 2479

    # 計算圖片最大寬高
    images_per_row = len(rows[0])
    images_per_col = 2
    min_spacing_w = 40
    min_spacing_h = 40
    middle_spacing = 200

    max_image_width = (total_width - min_spacing_w * (images_per_row+1)) // images_per_row
    max_image_height = (total_height - middle_spacing) // 2

    # 縮放圖片
    rows = [[scale_to_fit(cv.imread(fn), max_image_width, max_image_height) for fn in row] for row in rows]

    # 畫上圖片

    # 建立資料夾
    folder = 'result'
    try:
        os.makedirs(folder)
    except Exception:
        pass

    imagelist = []
    for i in range(len(rows)//2): # 0, 1, ... , 24
        canvas = np.zeros((total_height, total_width, 3), np.uint8)
        canvas[:] = (255, 255, 255)

        upper_images = rows[i*2]
        lower_images = rows[i*2+1]

        upper_padding = (total_width - np.sum([image.shape[1] for image in upper_images])) // (len(upper_images)+1)
        lower_padding = (total_width - np.sum([image.shape[1] for image in lower_images])) // (len(lower_images)+1)
        assert upper_padding >= min_spacing_w and lower_padding >= min_spacing_w

        # 評分 padding
        x_upper = upper_padding
        for index, image in enumerate(upper_images):
            y = 0

            draw_on(canvas, image, x_upper, y)
            x_upper += image.shape[1] + upper_padding

        x_lower = lower_padding
        for index, image in enumerate(lower_images):
            y = total_height - image.shape[0]

            draw_on(canvas, image, x_lower, y)
            x_lower += image.shape[1] + lower_padding

        # 產出圖片
        cv.imwrite(f'{folder}/{i}.jpg', canvas)
        imagelist.append(f'{folder}/{i}.jpg')

    pdf = FPDF()
    w = 297
    h = 210
    for image in imagelist:
        pdf.add_page("L")
        pdf.image(image,0,0,w,h)
    pdf.output("result.pdf", "F")