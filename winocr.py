import sys
import asyncio
import base64
import copy
import pprint

import winsdk
# pip3 install winsdk
from winsdk.windows.media.ocr import OcrEngine
from winsdk.windows.globalization import Language
from winsdk.windows.graphics.imaging import *
from winsdk.windows.security.cryptography import CryptographicBuffer
from PIL import Image


class rect:
    def __init__(self, x, y, w, h):
        self.x = x
        self.y = y
        self.width = w
        self.height = h

    def __repr__(self):
        return 'rect(%d, %d, %d, %d)' % (self.x, self.y, self.width, self.height)

    def right(self):
        return self.x + self.width

    def bottom(self):
        return self.y + self.height

    def set_right(self, value):
        self.width = value - self.x

    def set_bottom(self, value):
        self.height = value - self.y


def dump_rect(rtrect: winsdk.windows.foundation.Rect):
    return rect(rtrect.x, rtrect.y, rtrect.width, rtrect.height)


def dump_ocrword(word):
    return {
        'bounding_rect': dump_rect(word.bounding_rect),
        'text': word.text
    }


def merge_words(words):
    if len(words) == 0:
        return words
    new_words = [copy.deepcopy(words[0])]
    words = words[1:]
    for word in words:
        lastnewword = new_words[-1]
        lastnewwordrect = new_words[-1]['bounding_rect']
        wordrect = word['bounding_rect']
        if len(word['text']) == 1 and wordrect.x - lastnewwordrect.right() <= wordrect.width * 0.2:
            lastnewword['text'] += word['text']
            lastnewwordrect.x = min((wordrect.x, lastnewwordrect.x))
            lastnewwordrect.y = min((wordrect.y, lastnewwordrect.y))
            lastnewwordrect.set_right(max((wordrect.right(), lastnewwordrect.right())))
            lastnewwordrect.set_bottom(max((wordrect.bottom(), lastnewwordrect.bottom())))
        else:
            new_words.append(copy.deepcopy(word))
    return new_words


def dump_ocrline(line):
    words = list(map(dump_ocrword, line.words))
    merged = merge_words(words)
    return {
        'text': line.text,
        'words': words,
        'merged_words': merged,
        'merged_text': ' '.join(map(lambda x: x['text'], merged))
    }


def dump_ocrresult(ocrresult):
    lines = list(map(dump_ocrline, ocrresult.lines))
    return {
        'text': ocrresult.text,
        'text_angle': ocrresult.text_angle if ocrresult.text_angle else None,
        'lines': lines,
        'merged_text': ' '.join(map(lambda x: x['merged_text'], lines))
    }


def ibuffer(s):
    """create winsdk IBuffer instance from a bytes-like object"""
    return CryptographicBuffer.decode_from_base64_string(base64.b64encode(s).decode('ascii'))


def swbmp_from_pil_image(img):
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    pybuf = img.tobytes()
    rtbuf = ibuffer(pybuf)
    return SoftwareBitmap.create_copy_from_buffer(rtbuf, BitmapPixelFormat.RGBA8, img.width, img.height,
                                                  BitmapAlphaMode.STRAIGHT)


async def ensure_coroutine(awaitable):
    return await awaitable


def blocking_wait(awaitable):
    return asyncio.run(ensure_coroutine(awaitable))


def recognize_pil_image(img, lang):
    lang = Language(lang)
    assert (OcrEngine.is_language_supported(lang))
    eng = OcrEngine.try_create_from_language(lang)
    swbmp = swbmp_from_pil_image(img)
    return dump_ocrresult(blocking_wait(eng.recognize_async(swbmp)))


def recognize_file(filename, lang):
    img = Image.open(filename)
    return recognize_pil_image(img, lang)


if __name__ == '__main__':
    if 2 <= len(sys.argv) <= 3:
        lang = 'zh-hans-cn' if len(sys.argv) == 2 else sys.argv[1]
        result = recognize_file(sys.argv[-1], lang)
        pprint.pprint(result, width=128)
    else:
        print('usage: %s [language=zh-hans-cn] filename' % sys.argv[0])
        langs = list(map(lambda x: x.language_tag, OcrEngine.available_recognizer_languages))
        print('installed languages:', ', '.join(langs))
