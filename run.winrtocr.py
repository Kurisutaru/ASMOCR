import ctypes
import difflib
import sqlite3
import time
from io import BytesIO

import win32con
import winocr
from PIL import ImageGrab, ImageDraw, Image, ImageEnhance, ImageOps
from playsound import playsound
from win32api import SetCursorPos, mouse_event
from win32gui import EnumWindows, SetForegroundWindow, GetClientRect, GetWindowRect, GetWindowText

from PPOCR_api import GetOcrApi
import tbpu
from PPOCR_visualize import visualize

print("================================")
print("Source Code: https://github.com/paulzzh/ASMOCR")
print("Demo Video: https://www.bilibili.com/video/BV1tm4y1T7Dr")
print("Not for commercial use; resale is strictly prohibited")
print("================================")
print("Please start the DMM client before running the script")
print("It will automatically click when you enter the mini-game")
print("================================")

offset = 64
alert_sound_file_name = "alert.mp3"

debug_window = False
debug_ocr = False

window_title = "PrincessConnectReDive"

sub = (1156, 584)
retry = (950, 655)

question = (32, 167 + 10, 32 + 678, 167 + 155)

top_left = (32, 425, 32 + 320, 425 + 100)
top_right = (395, 425, 395 + 320, 425 + 100)
bottom_left = (32, 575, 32 + 320, 575 + 100)
bottom_right = (395, 575, 395 + 320, 575 + 100)

multi_answer_top_left = (32 + int(offset / 2), 425, 32 + 320, 425 + 100)
multi_answer_top_right = (395 + int(offset / 2), 425, 395 + 320, 425 + 100)
multi_answer_bottom_left = (32 + int(offset / 2), 575, 32 + 320, 575 + 100)
multi_answer_bottom_right = (395 + int(offset / 2), 575, 395 + 320, 575 + 100)

top_left_click_coords = ((top_left[0] + top_left[2]) // 2, (top_left[1] + top_left[3]) // 2)
top_right_click_coords = ((top_right[0] + top_right[2]) // 2, (top_right[1] + top_right[3]) // 2)
bottom_left_click_coords = ((bottom_left[0] + bottom_left[2]) // 2, (bottom_left[1] + bottom_left[3]) // 2)
bottom_right_click_coords = ((bottom_right[0] + bottom_right[2]) // 2, (bottom_right[1] + bottom_right[3]) // 2)
click_coords_array = [top_left_click_coords, top_right_click_coords, bottom_left_click_coords, bottom_right_click_coords]

yes = (top_left_click_coords[0], (top_left_click_coords[1] + bottom_left_click_coords[1]) // 2)
no = (top_right_click_coords[0], (top_right_click_coords[1] + bottom_right_click_coords[1]) // 2)

print("Window Title:", window_title)
print("Submit Button:", sub)
print("Retry Button:", retry)
print("Question Area:", question)
print("Top Left Button:", top_left, top_left_click_coords)
print("Top Right Button:", top_right, top_right_click_coords)
print("Bottom Left Button:", bottom_left, bottom_left_click_coords)
print("Bottom Right Button:", bottom_right, bottom_right_click_coords)
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
        draw.rectangle(top_left, outline=(255, 0, 0))
        draw.text((top_left_click_coords[0], top_left_click_coords[1]), "topleft", "red")
        draw.rectangle(top_right, outline=(255, 0, 0))
        draw.text((top_right_click_coords[0], top_right_click_coords[1]), "topright", "red")
        draw.rectangle(bottom_left, outline=(255, 0, 0))
        draw.text((bottom_left_click_coords[0], bottom_left_click_coords[1]), "bottomleft", "red")
        draw.rectangle(bottom_right, outline=(255, 0, 0))
        draw.text((bottom_right_click_coords[0], bottom_right_click_coords[1]), "bottomright", "red")

        draw.text((yes[0], yes[1]), "yes", "red")
        draw.text((no[0], no[1]), "no", "red")
        draw.text((sub[0], sub[1]), "sub", "red")
        draw.text((retry[0], retry[1]), "retry", "red")

        frame.save(window_title + ".png", "PNG")

    def screenshot(self):
        SetForegroundWindow(pcrd_window)
        x1, y1, x2, y2 = GetClientRect(self.hwnd)
        self.width = x2 - x1
        self.height = y2 - y1

        wx1, wy1, wx2, wy2 = GetWindowRect(self.hwnd)
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
        SetForegroundWindow(pcrd_window)
        x = self.xbase + pos[0]
        y = self.ybase + pos[1]
        SetCursorPos([x, y])
        mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y)
        SetCursorPos([x, y])
        mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y)
        time.sleep(0.333)


class OCR:
    def __init__(self, argument=None):
        if argument is None:
            argument = {'config_path': "models/config_japan.txt"}
        self.ocr = GetOcrApi(r".\PaddleOCR-json\PaddleOCR-json.exe", argument)

    def get_single_line_from_ocr_result(self, frame_data, use_winrt_ocr: bool = False, pre_process: bool = False,
                                        inverse: bool = False):
        buffered = BytesIO()
        final_image = frame_data

        if pre_process:
            # new_width = int(original_width * 1.5)
            # new_height = int(original_height * 1.5)
            # original_width, original_height = frame_data.size
            # resized_image = frame_data.resize((new_width, new_height)).convert('L')
            contrast = ImageEnhance.Contrast(frame_data.convert('L'))
            final_image = contrast.enhance(2)

        if inverse:
            final_image = ImageOps.invert(final_image)

        final_image.save(buffered, format="PNG")

        if debug_window:
            final_image.save("123.png", format="PNG")

        if use_winrt_ocr:
            result = winocr.recognize_pil_image(final_image, 'ja')
        else:
            result = self.ocr.runBytes(buffered.getvalue())

        if debug_ocr:
            with open('123.png', 'rb') as image_file:
                # Read the contents of the image file as bytes
                image_bytes = image_file.read()
            result = self.ocr.runBytes(image_bytes)

        if use_winrt_ocr:
            if len(result['lines']) > 0:
                # Exclude list of word that fill in the upper answer box
                # exclude_words = ["かな", "がな", "カナ", "ガナ", "ガす", "漢字", "数字", "漢宮"]
                # filtered_data = [item for item in result["lines"] if item["merged_text"] not in exclude_words]
                return self.replace_japanese_symbol(result['merged_text'].replace(' ', ''))
            else:
                return ""
        else:
            if result['code'] == 100:
                # Exclude list of word that fill in the upper answer box
                exclude_words = ["かな", "がな", "カナ", "ガナ", "ガす", "漢字", "数字", "漢宮"]
                # filtered_data = [item for item in result["data"] if not any(word in item["text"] for word in exclude_words)]
                # filtered_data = [item for item in result["data"] if not any(exclude_word in item["text"] for exclude_word in exclude_words)]
                filtered_data = [item for item in result["data"] if item["text"] not in exclude_words]
                txts = tbpu.run_merge_line_h_m_paragraph(filtered_data)
                if debug_window:
                    img1 = visualize(result["data"], "123.png").get(isOrder=True)
                    img2 = visualize(txts, "123.png").get(isOrder=True)
                    visualize.createContrast(img1, img2).show()
                if len(txts) == 1:
                    text = txts[0]['text'] if txts and 'text' in txts[0] else ''
                else:
                    text = ''.join(item['text'] for item in txts)
                return self.replace_japanese_symbol(text)
            else:
                return ""

    @staticmethod
    def replace_japanese_symbol(input_text: str):
        list_of_japanese_symbol_counterpart = [
            ('-', 'ー'), ('+', '＋'), (',', '、'), ('!', '！'), ('?', '？'),
            ("'", '’'), ('"', '”'), ('(', '（'), (')', '）')
        ]

        for symbol, counterpart in list_of_japanese_symbol_counterpart:
            text = input_text.replace(symbol, counterpart)

        return text


def set_dpi_awareness():
    awareness = ctypes.c_int()
    error_code = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
    error_code = ctypes.windll.shcore.SetProcessDpiAwareness(2)
    success = ctypes.windll.user32.SetProcessDPIAware()


def get_windows_by_title(title_text):
    def _window_callback(_hwnd, all_windows):
        all_windows.append((_hwnd, GetWindowText(_hwnd)))

    windows = []
    EnumWindows(_window_callback, windows)
    return [(hwnd, title) for hwnd, title in windows if title_text in title]


try:
    set_dpi_awareness()
    wins = get_windows_by_title(window_title)
    print("Window Handle:", wins)
    print("Defaulting to the first one; modify the configuration file if it's not correct")
    pcrd_window = wins[0][0]
    SetForegroundWindow(pcrd_window)
    time.sleep(1)
    window = Window(pcrd_window)
    ocr_class = OCR()
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


asm_data = query_jp_db(
    "select * from asm_data ORDER BY asm_id ASC")
asm_true_or_false_data = query_jp_db("select * from asm_true_or_false_data")
asm_4_choice_data = query_jp_db("select * from asm_4_choice_data")
asm_many_answers_data = query_jp_db("select * from asm_many_answers_data")

asm_true_or_false_data_dict = {x["asm_id"]: x for x in asm_true_or_false_data}
asm_4_choice_data_dict = {x["asm_id"]: x for x in asm_4_choice_data}
asm_many_answers_data_dict = {x["asm_id"]: x for x in asm_many_answers_data}

# Find the most similar question
ques = None


def diff_asm(obj):
    try:
        return difflib.SequenceMatcher(None, obj["detail"], ques).quick_ratio()
    except:
        return 0.0


def diff_ab(a, b):
    return difflib.SequenceMatcher(None, a, b).quick_ratio()


asm_id = None
wait = 0
click_count = None

exclude_asm_id = [2630002]
do_not_click = False
while True:
    click_count = 0
    frame = window.screenshot()

    quesimg = frame.crop(question)  # Crop to the question area

    ques = ocr_class.get_single_line_from_ocr_result(quesimg, use_winrt_ocr=True)
    data = max(asm_data, key=diff_asm, default='')
    score = diff_asm(data)
    if asm_id == data["asm_id"] or score < 0.5:  # Repeat or low match rate
        if wait >= 3:  # Maybe it's settled
            window.click(retry)
            wait = 0
        time.sleep(2)
        wait += 1
        continue

    asm_id = data["asm_id"]

    if asm_id in exclude_asm_id:
        do_not_click = True

    wait = 0
    print("================================")
    print("ASM:", asm_id)
    print("OCR:", ques)
    print("QUE:", "{:.2%}".format(score), data["detail"])

    time.sleep(0.15)  # It seems like recognition is too fast, clicking too fast might not work

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
        if do_not_click:
            playsound(alert_sound_file_name)
            do_not_click = False
            time.sleep(5)
        else:
            img1 = frame.crop(top_left)  # Top left
            img2 = frame.crop(top_right)  # Top right
            img3 = frame.crop(bottom_left)  # Bottom left
            img4 = frame.crop(bottom_right)  # Bottom right
            ansimgs = [img1, img2, img3, img4]
            for i in range(4):
                an = ocr_class.get_single_line_from_ocr_result(frame_data=ansimgs[i])
                ans_ocr.append(an)
                if an == ans_str:
                    score = 1.0
                    best = i
                    break
                if diff_ab(an, ans_str) > score:
                    score = diff_ab(an, ans_str)
                    best = i

            print("OCR:", "{:.2%}".format(score), ans_ocr[best])
            if score >= 0.8:
                window.click(click_coords_array[best])
                click_count += 1

            if click_count == 1:
                window.click(sub)
            else:
                playsound(alert_sound_file_name)
    else:  # Multiple choice
        ans = asm_many_answers_data_dict[asm_id]
        ans_strs = []
        for i in ["1", "2", "3", "4"]:
            if ans["is_correct_" + i] == 1:
                ans_strs.append(ans["choice_" + i])
                print("ANS:", ans["choice_" + i])
        if do_not_click:
            playsound(alert_sound_file_name)
            do_not_click = False
            time.sleep(5)
        else:
            img1 = frame.crop(multi_answer_top_left)  # Top left
            img2 = frame.crop(multi_answer_top_right)  # Top right
            img3 = frame.crop(multi_answer_bottom_left)  # Bottom left
            img4 = frame.crop(multi_answer_bottom_right)  # Bottom right
            ansimgs = [img1, img2, img3, img4]
            ans_ocr = []
            for i in range(4):
                an = ocr_class.get_single_line_from_ocr_result(frame_data=ansimgs[i])
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
            for i in range(len(scores)):
                if scores[i] >= 0.8:
                    window.click(click_coords_array[bests[i]])
                    click_count += 1

            if click_count == len(bests):
                window.click(sub)
            else:
                playsound(alert_sound_file_name)

    print("================================")
