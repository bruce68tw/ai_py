"""
pip install python-docx
用途：將 .docx（以法規或工作規則等格式為主）解析成標準化 JSON。
輸出格式：一個 list，元素為每個「條」或「次標」的 dict，範例如下：
  {
    "article0": "第1條 條目完整標題",
    "article": "條目標題（不含第X條字樣）",
    "page": None,      # 預留欄位（python-docx 無法取得頁碼）
    "body": "條文內容（已統一清理換行）"
  }
說明：
- 解析採逐段落掃描，透過正則判斷章、條、次標與條列項目。
- 章只作為標題補全用途（不直接輸出章物件）。
- 支援次標（a., b. 等）作為條下獨立輸出項目。
- 條列（如「一、二、三」）會保留項目前綴並以換行分隔；續行會接回同一項目。
保留設計：
- page 欄位保留為 None，未來若以其他方式取得頁碼可在 extract_page() 擴充。
"""

import os
import re
import json
import pdb
import sys
from docx import Document
#import logging

#logging.basicConfig(level=logging.INFO)

# 常數
Deleted = "(刪除)"  #刪除標記
Article0 = "article0"   #欄位
Article = "article"     #欄位
Body = "body"           #欄位
Carrier = "\n"  #換行符號

# 傳入參數: word檔名(不含副檔名)
WordFname = ""

# 正則陣列, 內容由傳入參數決定
Regs = [];

L0Tail = '章'

# === 正則定義 ===
# 階層名稱常數
L0Tail = '章'
L1Tail = '條'
L2Tail = '次標'

CN_NUM = '[一二三四五六七八九十]+'
#todo: word結構層級以參數傳入
def getRe(reg: str):
    return re.compile(
        rf'^(第\s*({CN_NUM})\s*{reg})\s*(.*)'
    )

# 正則
# r'...' 裡面的 \ 不會被 Python 當跳脫字元
# ^ 行首錨點, 強制從一行最前面開始比對
# \s* 0個(含)以上空白字元
# [...]：字元集合（選一）
# +：前面集合出現 一次以上
# (.*) 第二個捕捉群組, .：任意字元（不含換行）,*：0 次以上, 抓「該行剩下的所有內容」

# Get(取值), Chk(判斷)
L0TitleGetRe = re.compile(r'^(第\s*[一二三四五六七八九十]+\s*章)\s*(.*)')   #group(1)傳回第x章
L1TitleGetRe = re.compile(r'^(第\s*[一二三四五六七八九十]+\s*條)\s*(.*)')   #第x條
L1ItemGetRe = re.compile(r'^([一二三四五六七八九十]+)[、\.]\s*(.*)')        #ex: 一、
L2ItemChkRe = re.compile(r'^\([一二三四五六七八九十]+\)\s*')                #ex: (一)
#L1ItemChkRe = re.compile(r'^[一二三四五六七八九十]+[、\.]\s*')              #ex: 一、

# instance variables
_l0Article0 = ""
_l0Article = ""
_l1Item = {}
_l2Items = []
_isL2 = False

def getPage(paragraph):
    """
    python-docx 無法取得最終排版後的頁碼（頁碼屬於 layout 結果），
    回傳：整數頁碼或 None
    """
    return None

#有2個title時使用@@分隔
def getTitle0(title1, title2):
    return title1 + '@@' + title2

def articleRow(title0, article0, article, page, body):
    return {
        "title0": title0,
        "article0": article0,
        "article": article,
        "page": page,
        "body": body
    }

def removeCarrier(text):
    return text.replace('\n', '')

def getText(carrier, text):
    text = removeCarrier(text)
    if carrier:
        text = Carrier + text
    return text

# add L1 row to results[]
def addL1(results):
    global _l1Item, _isL2, _l2Items
    if not _l1Item:
        return
    
    results.append(_l1Item)
    _l1Item = {}

    addL2(results)

def addL1Body(carrier, text):
    global _l1Item

    #debug
    #pdb.set_trace()

    _l1Item[Body] += getText(carrier, text)

#無條件reset L2
def resetL2():
    global _isL2, _l2Items
    #if not _isL2:
    #    return
    
    _isL2 = False
    _l2Items = []

def addL2(results):
    global _isL2, _l2Items

    #append l2 items if exist
    if _isL2:
        for item in _l2Items:
            results.append(item)            
    #reset
    resetL2()

def addL2Body(carrier, text):
    global _l2Items
    #debug
    #pdb.set_trace()

    _l2Items[-1][Body] += getText(carrier, text)     #-1表示最後一個陣列元素

def isDeleted(text):
    return (Deleted in text)

def wordToJson(wordPath, outputPath):
    """
    將單一 .docx 轉成 outputPath/JSON 檔案 
    解析策略（階層）：
    - 第0階：章（只暫存章標題作為補全使用，不輸出章物件）
    - 第1階：條（以 "第...條" 為開頭建立新的輸出物件）
    - 第2階：次標（a,b,c）視為獨立輸出物件，屬於某條之下
    - 條列（如「一、二、三」）視為同一條的條列內容，會以換行保留每項
    """

    global _l0Article0, _l0Article, _l1Item, _l2Items, _isL2
    doc = Document(wordPath)
    results = []  # 最終輸出之 list

    #loop 讀取 word 檔
    for para in doc.paragraphs:
        paraText = para.text.strip()

        # 空段落跳過
        if not paraText or isDeleted(paraText):
            continue

        # === 第0階 title ===
        titleRe = L0TitleGetRe.match(paraText)
        if titleRe:
            #handle previous
            addL1(results)

            #handle this
            _l0Article0 = titleRe.group(0).strip()
            _l0Article = titleRe.group(2).strip()
            continue

        # === 第1階 title ===
        titleRe = L1TitleGetRe.match(paraText)
        if titleRe:
            #handle previous
            addL1(results)

            #handle this, 暫時記錄在 _l1Item
            article0 = titleRe.group(0).strip()
            article = titleRe.group(2).strip()
            #article前面加上第0階主題if need (for embedding vector)
            #if not (article in _l0Article or _l0Article in article):
            #    article = f"{_l0Article}-{article}"

            _l1Item = articleRow(_l0Article0, article0, article, getPage(para), "")
            continue

        # === 第1階 item, ex: 一、 ===
        itemRe = L1ItemGetRe.match(paraText)
        if itemRe:
            #handle previous
            addL2(results)

            #handle this
            addL1Body(True, paraText)

            #此項有可能包含L2 item, 所以先對 _l2Items 初始化, 但是不設定 _isL2
            title0 = getTitle0(_l0Article0, _l1Item[Article0])
            article = itemRe.group(2).strip()            
            _l2Items.append(articleRow(title0, paraText, article, getPage(para), ""))
            continue

        # === 第2階 item, ex: (一) ===
        #itemRe = L2ItemChkRe.match(paraText)
        if L2ItemChkRe.match(paraText):
            #handle previous

            #addL2(results)

            #它同時也是 l1 item, 所以加入 l1 body
            addL1Body(True, paraText)

            #set variables
            _isL2 = True

            #handle this, add L2 row with body
            #title0 = getTitle0(_l0Article0, _l1Item[Article0])
            #_l2Items.append(articleRow(title0, _l1Item[Article0], _l1Item[Article], getPage(para), paraText));

            addL2Body(True, paraText)

            #此項有可能包含L2 item, 所以先對 _l2Items 初始化, 但是不設定 _isL2
            #title0 = getTitle0(_l0Article0, _l1Item[Article0])
            #article = itemRe.group(2).strip()            
            #_l2Items.append(articleRow(title0, paraText, article, getPage(para), ""))

            continue

        # === 其他情形 ===
        #if paraText.strip() == '四、特別休假':
        #if _l1Item and "特別休假" in _l1Item[Article]:
            #debug
        #    pdb.set_trace()

        #if _l1Item:
            #print(_l1Item[Article])

        #debug
        #pdb.set_trace()

        #已經產生 _L1Item 
        if not _l1Item:
            continue

        #handle this
        addL1Body(False, paraText)

        #如果是L2, 則加入L2 body(含換行符號)
        if _isL2:
            addL2Body(True, paraText)
    #exit for loop

    #寫入最後一筆
    addL1(results)

    # 寫入 JSON檔，使用 ensure_ascii=False 以保留中文
    with open(outputPath, "w", encoding="utf-8") as file:
        json.dump(results, file, ensure_ascii=False, indent=2)


# === 主程式 ===
# 傳入3個參數: python檔名, word檔名(不含副檔名), 層級陣列
# 層級陣列的中間參數可為以下內容:
#   C1:中文數字, ex: 一二三四五六七八九十
#   N1:數字(半形), ex: 12345678910
#   N2:數字(全形), ex: １２３４５６７８９１０
if __name__ == "__main__":
    # check 參數數量必須為3
    if len(sys.argv) != 3:
        print("用法: python xxx.py word檔名(不含副檔名) 層級陣列")
        sys.exit(1)

    # check 參數2: 層級陣列必須是3的倍數
    regArgs = sys.argv[2]
    if len(regArgs) % 3 != 0:
        print("參數3(層級陣列)必須是3的倍數!!")
        sys.exit(1)
    
    # set instance variables
    WordFname = sys.argv[1]
    
    # set Regs array
    for i in range(0, len(regArgs), 3):
        left, mid, right = regArgs[i:i+3]
        regex = '^'
        if left:
            regex += left
        if mid:
            regex += mid
        if right:
            regex += right

        # 擷取結構後的內容
        regex += r'\s*(.*)$'

        patterns.append(re.compile(regex))

    # 範例執行：開發時可改為從命令列參數讀入路徑
    inputPath = f"input/{WordFname}.docx"
    outputDir = "output"

    if not os.path.isfile(inputPath):
        raise FileNotFoundError(f"找不到檔案：{inputPath}")

    os.makedirs(outputDir, exist_ok=True)

    fname = os.path.basename(inputPath)
    outputPath = os.path.join(
        outputDir,
        fname.replace(".docx", ".json")
    )

    wordToJson(inputPath, outputPath)
    print(f"Converted: {fname}")
