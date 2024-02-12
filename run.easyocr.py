import time, sqlite3, difflib, ctypes

import rapidfuzz
import win32api, win32gui, win32con
from PIL import Image, ImageGrab, ImageDraw, ImageOps, ImageEnhance
from io import BytesIO

from easyocr import easyocr
from rapidfuzz import process

from PPOCR_api import GetOcrApi
from PPOCR_visualize import visualize
import tbpu

print("================================")
print("Princess Connect! Re:Dive - 2023 Academy Quiz Solver")
print("================================")

debug_window = False
debug_ocr = False

offset = 48

window_title = "PrincessConnectReDive"
sub = (1156, 584)
retry = (950, 655)

question = (32, 167 + 13, 32 + 678, 167 + 163)

topleft = (32 + offset, 425, 32 + 320, 425 + 100)
topright = (395 + offset, 425, 395 + 320, 425 + 100)
bottomleft = (32 + offset, 575, 32 + 320, 575 + 100)
bottomright = (395 + offset, 575, 395 + 320, 575 + 100)

b1 = ((topleft[0] + topleft[2] + offset) // 2, (topleft[1] + topleft[3]) // 2)
b2 = ((topright[0] + topright[2] + offset) // 2, (topright[1] + topright[3]) // 2)
b3 = ((bottomleft[0] + bottomleft[2] + offset) // 2, (bottomleft[1] + bottomleft[3]) // 2)
b4 = ((bottomright[0] + bottomright[2] + offset) // 2, (bottomright[1] + bottomright[3]) // 2)
bs = [b1, b2, b3, b4]

yes = (b1[0], (b1[1] + b3[1]) // 2)
no = (b2[0], (b2[1] + b4[1]) // 2)

print("Window Title:", window_title)
print("Submit Button:", sub)
print("Retry Button:", retry)
print("Question Area:", question)
print("Top Left Button:", topleft, b1)
print("Top Right Button:", topright, b2)
print("Bottom Left Button:", bottomleft, b3)
print("Bottom Right Button:", bottomright, b4)
print("Correct:", yes)
print("Incorrect:", no)


class Window:
    def __init__(self, hwnd):
        self.ybase = None
        self.xbase = None
        self.height = None
        self.width = None
        if not hwnd:
            return
        self.hwnd = hwnd
        frame = self.screenshot()
        print("Resolution:", frame.size)
        print("Base Coordinates:", (self.xbase, self.ybase))
        draw = ImageDraw.Draw(frame)
        draw.rectangle(question, outline=(255, 0, 0))
        draw.text((question[0], question[1]), "question", "red")
        draw.rectangle(topleft, outline=(255, 0, 0))
        draw.text((b1[0], b1[1]), "topleft", "red")
        draw.rectangle(topright, outline=(255, 0, 0))
        draw.text((b2[0], b2[1]), "topright", "red")
        draw.rectangle(bottomleft, outline=(255, 0, 0))
        draw.text((b3[0], b3[1]), "bottomleft", "red")
        draw.rectangle(bottomright, outline=(255, 0, 0))
        draw.text((b4[0], b4[1]), "bottomright", "red")

        draw.text((yes[0], yes[1]), "yes", "red")
        draw.text((no[0], no[1]), "no", "red")
        draw.text((sub[0], sub[1]), "sub", "red")
        draw.text((retry[0], retry[1]), "retry", "red")

        frame.save(window_title + ".png", "PNG")

    def screenshot(self):
        win32gui.SetForegroundWindow(pcrd_window)
        x1, y1, x2, y2 = win32gui.GetClientRect(self.hwnd)
        self.width = x2 - x1
        self.height = y2 - y1

        wx1, wy1, wx2, wy2 = win32gui.GetWindowRect(self.hwnd)
        bx = wx1
        by = wy1
        # normalize to origin
        wx1, wx2 = wx1 - wx1, wx2 - wx1
        wy1, wy2 = wy1 - wy1, wy2 - wy1
        # compute border width and title height
        bw = int((wx2 - x2) / 2.)
        th = wy2 - y2 - bw
        # calc offset x and y taking into account border and titlebar, screen coordiates of client rect
        sx = bw
        sy = th

        self.xbase = bx + sx
        self.ybase = by + sy

        left, top = self.xbase, self.ybase
        right, bottom = left + self.width, top + self.height

        _frame = ImageGrab.grab(bbox=(left, top, right, bottom))

        if debug_window:
            image = Image.open("frame.png")
            return image

        return _frame

    def click(self, pos):
        win32gui.SetForegroundWindow(pcrd_window)
        x = self.xbase + pos[0]
        y = self.ybase + pos[1]
        win32api.SetCursorPos([x, y])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y)
        win32api.SetCursorPos([x, y])
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y)
        time.sleep(0.5)


class OCR:
    def __init__(self, argument=None):
        # if argument is None:
        #     argument = {'config_path': "models/config_japan.txt"}
        self.ocr = GetOcrApi(r".\PaddleOCR-json\PaddleOCR-json.exe", argument)
        self.easy_ocr = easyocr.Reader(['ja', 'en'], detector='dbnet18')

    def get_single_line_from_ocr_result(self, frame_data, resize: bool = False, inverse: bool = False):
        buffered = BytesIO()
        original_width, original_height = frame_data.size
        if resize:
            new_width = int(original_width * 2)
            new_height = int(original_height * 2)
            resized_image = frame_data.resize((new_width, new_height))
        #enhancer = ImageEnhance.Contrast(frame_data)
        #final_image = enhancer.enhance(2)
        final_image = resized_image

        if inverse:
            final_image = ImageOps.invert(final_image)

        final_image.save(buffered, format="PNG")

        if debug_window:
            final_image.save("123.png", format="PNG")

        #result = self.ocr.runBytes(buffered.getvalue())
        result = self.easy_ocr.readtext(buffered.getvalue(), batch_size=5, paragraph=True)

        if debug_ocr:
            with open('123.png', 'rb') as image_file:
                # Read the contents of the image file as bytes
                image_bytes = image_file.read()
            # result = self.ocr.runBytes(image_bytes)
            result = self.easy_ocr.readtext(image_bytes)

        #if result['code'] == 100:
        # Exclude list of word that fill in the upper answer box
        exclude_words = ["かな", "がな", "カナ", "ガナ", "ガす", "漢字", "数字", "漢宮"]
        # filtered_data = [item for item in result["data"] if not any(word in item["text"] for word in exclude_words)]
        # filtered_data = [item for item in result["data"] if not any(exclude_word in item["text"] for exclude_word in exclude_words)]
        # filtered_data = [item for item in result["data"] if item["text"] not in exclude_words]
        filtered_data = [item for item in result if item[1] not in exclude_words]
        #txts = tbpu.run_merge_line_h_m_paragraph(filtered_data)
        # if debug_window:
        #     print("")
        #     img1 = visualize(result["data"], "123.png").get(isOrder=True)
        #     img2 = visualize(txts, "123.png").get(isOrder=True)
        #     visualize.createContrast(img1, img2).show()
        #if len(txts) == 1:
        #    text = txts[0]['text'] if txts and 'text' in txts[0] else ''
        #else:
        #    text = ''.join(item['text'] for item in txts)

        text = ''.join(item[1] for item in filtered_data)

        return text

        #else:
        #    return ""


def set_dpi_awareness():
    awareness = ctypes.c_int()
    error_code = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
    error_code = ctypes.windll.shcore.SetProcessDpiAwareness(2)
    success = ctypes.windll.user32.SetProcessDPIAware()


def get_windows_by_title(title_text):
    def _window_callback(_hwnd, all_windows):
        all_windows.append((_hwnd, win32gui.GetWindowText(_hwnd)))

    windows = []
    win32gui.EnumWindows(_window_callback, windows)
    return [(hwnd, title) for hwnd, title in windows if title_text in title]


try:
    set_dpi_awareness()
    wins = get_windows_by_title(window_title)
    print("Window Handle:", wins)
    print("Defaulting to the first one; modify the configuration file if it's not correct")
    pcrd_window = wins[0][0]
    win32gui.SetForegroundWindow(pcrd_window)
    time.sleep(1)
    window = Window(pcrd_window)
    ocr = OCR()
except Exception as e:
    print(f"Caught an exception: {e}")
    print("Initializing failed, exiting program")
    exit(1)


# Read mini-game data
def query_jp_db(query, args=(), one=False):
    # https://stackoverflow.com/questions/3286525/return-sql-table-as-json-in-python
    conn = sqlite3.connect('redive_jp.db')
    cur = conn.cursor()
    cur.execute(query, args)
    r = [dict((cur.description[i][0], value) \
              for i, value in enumerate(row)) for row in cur.fetchall()]
    cur.close()
    conn.close()
    return (r[0] if r else None) if one else r


exclude_character = "、「」？・ー！。"

asm_data = query_jp_db(
    "select * from (select * from asm_data ORDER BY asm_id DESC) GROUP BY detail ORDER BY asm_id ASC")
asm_true_or_false_data = query_jp_db("select * from asm_true_or_false_data")
asm_4_choice_data = query_jp_db("select * from asm_4_choice_data")
asm_many_answers_data = query_jp_db("select * from asm_many_answers_data")

asm_true_or_false_data_dict = {x["asm_id"]: x for x in asm_true_or_false_data}
asm_4_choice_data_dict = {x["asm_id"]: x for x in asm_4_choice_data}
asm_many_answers_data_dict = {x["asm_id"]: x for x in asm_many_answers_data}

# Find the most similar question
question_ocr = None
question_data = [item['detail'] for item in asm_data]


def diff_asm(obj):
    try:
        return difflib.SequenceMatcher(None, obj["detail"], question_ocr).quick_ratio()
    except:
        return 0.0


def diff_ab(a, b):
    return difflib.SequenceMatcher(None, a, b).quick_ratio()


asm_id = None
wait = 0
click_count = None
while True:
    click_count = 0
    frame = window.screenshot()
    question_img = frame.crop(question)  # Crop to the question area
    question_ocr = ocr.get_single_line_from_ocr_result(frame_data=question_img, inverse=True)
    # data = max(asm_data, key=diff_asm, default='')
    data_fuzzy_select = process.extractOne(query=question_ocr,
                                           choices=question_data,
                                           scorer=rapidfuzz.fuzz.QRatio,
                                           processor=rapidfuzz.utils.default_process)

    data = [item for item in asm_data if item["detail"] == data_fuzzy_select[0]][0]
    score = data_fuzzy_select[1] / 100 if data_fuzzy_select else 0
    if asm_id == data["asm_id"] or score < 0.7:  # Repeat or low match rate
        if wait >= 3:  # Maybe it's settled
            window.click(retry)
            wait = 0
        time.sleep(2)
        wait += 1
        continue

    wait = 0
    img1 = frame.crop(topleft)  # Top left
    img2 = frame.crop(topright)  # Top right
    img3 = frame.crop(bottomleft)  # Bottom left
    img4 = frame.crop(bottomright)  # Bottom right
    ansimgs = [img1, img2, img3, img4]

    print("================================")
    print("OCR:", question_ocr)
    print("QUE:", "{:.2%}".format(score), data["detail"])

    time.sleep(0.5)  # It seems like recognition is too fast, clicking too fast might not work

    asm_id = data["asm_id"]
    if asm_id < 2000000:  # True or False
        ans = asm_true_or_false_data_dict[asm_id]
        print("ANS:", ans["correct_answer"] == 1)
        if ans["correct_answer"] == 1:
            window.click(yes)
        else:
            window.click(no)

        window.click(sub)

    elif asm_id < 3000000:  # Single choice
        ans = asm_4_choice_data_dict[asm_id]
        ans_str = ans["choice_" + str(ans["correct_answer"])]
        print("ANS:", ans_str)

        score = 0.0
        best = 0
        ans_ocr = []
        for i in range(4):
            an = ocr.get_single_line_from_ocr_result(ansimgs[i])
            ans_ocr.append(an)
            if an == ans_str:
                score = 1.0
                best = i
                break
            if diff_ab(an, ans_str) > score:
                score = diff_ab(an, ans_str)
                best = i

        print("OCR:", "{:.2%}".format(score), ans_ocr[best])
        if score >= 0.65:
            window.click(bs[best])
            click_count += 1

        if click_count == 1:
            window.click(sub)

    else:  # Multiple choice
        ans = asm_many_answers_data_dict[asm_id]
        ans_strs = []
        for i in ["1", "2", "3", "4"]:
            if ans["is_correct_" + i] == 1:
                ans_strs.append(ans["choice_" + i])
                print("ANS:", ans["choice_" + i])

        ans_ocr = []
        for i in range(4):
            an = ocr.get_single_line_from_ocr_result(ansimgs[i])
            ans_ocr.append(an)
        scores = []
        bests = []
        for ans_str in ans_strs:
            score = 0.0
            best = 0
            for i in range(4):
                if diff_ab(ans_ocr[i], ans_str) > score:
                    score = diff_ab(ans_ocr[i], ans_str)
                    best = i
            scores.append(score)
            bests.append(best)

        for i in range(len(scores)):
            print("OCR:", "{:.2%}".format(scores[i]), ans_ocr[bests[i]])

        if len(bests) == len(ans_strs):
            for i in range(len(scores)):
                if scores[i] >= 0.65:
                    window.click(bs[bests[i]])
                    click_count += 1

        if click_count == len(ans_strs):
            window.click(sub)

    print("================================")
