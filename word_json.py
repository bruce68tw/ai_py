# filepath: d:\_project\ai_py\word_json.py
"""
word_json.py
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

注意：本檔案只加入註解，主要邏輯保持原樣。
"""

import os
import re
import json
from docx import Document
import pdb

# 常數
Deleted = "(刪除)"  #刪除標記
Article0 = "article0"
Article = "article"
Body = "body"

# === 正則定義 ===
# 階層名稱常數
L0Tail = '章'
L1Tail = '條'
L2Tail = '次標'

# 正則模式常數（可在需要時重用或調整）
# 判斷 & 取值
# r'...' 裡面的 \ 不會被 Python 當跳脫字元
# ^ 行首錨點, 強制從一行最前面開始比對
# \s* 0個(含)以上空白字元
# [...]：字元集合（選一）
# +：前面集合出現 一次以上
# (.*) 第二個捕捉群組, .：任意字元（不含換行）,*：0 次以上, 抓「該行剩下的所有內容」
L0TitlePtnGet = r'^第\s*([一二三四五六七八九十]+)\s*章\s*(.*)'
L1TitlePtnGet = r'^第\s*([一二三四五六七八九十]+)\s*條\s*(.*)'
#L2Ptn = r'^([a-z])[\.\t、]?\s*(.*)'
# 條列前綴（如「一、」）之前綴正則，供去除前綴使用
# 1階條文, ex: 一、
L1ItemPtnGet = r'^([一二三四五六七八九十]+)[、\.]\s*(.*)'
L1ItemPtnChk = r'^[一二三四五六七八九十]+[、\.]\s*'
# 偵測括號編號形式的子項（如 (一) 或 （一）），視為第2階主題起始
# 2階條文, ex: (一)
L2ItemPtnGet = r'^\(([一二三四五六七八九十]+)\)\s*(.*)'
L2ItemPtnChk = r'^\([一二三四五六七八九十]+\)\s*'
# PAT_PAREN_LINE = r'^\s*\(([一二三四五六七八九十]+)\)\s*(.*)'
#PAT_PAREN_HEADER_ONLY = r'^\s*\([一二三四五六七八九十]+\)\s*$'

# 其他通用正則常數（列表與括號標記）
# 判斷
#PAT_ARTICLE0_LIST = r'^[一二三四五六七八九十]+[、\.]\s*'

# 編譯正則
L0TitleGetRe = re.compile(L0TitlePtnGet)
L1TitleGetRe = re.compile(L1TitlePtnGet)
#L2Re = re.compile(L2Ptn)
L1ItemGetRe = re.compile(L1ItemPtnGet)
L1ItemChkRe = re.compile(L1ItemPtnChk)
L2ItemGetRe = re.compile(L2ItemPtnGet)
L2ItemChkRe = re.compile(L2ItemPtnChk)

# instance variables
_isL2 = False
_l2Texts = []

def resetL2():
    if _isL2:
        _l2Texts = []

    _isL2 = False

def extractPage(paragraph):
    """
    擷取頁碼（保留介面）

    python-docx 無法取得最終排版後的頁碼（頁碼屬於 layout 結果），
    因此此函式目前回傳 None。若未來改用其他工具或方法取得頁碼，
    可在此函式內實作而不需改動其他邏輯。

    參數：paragraph 為 docx 的 Paragraph 物件（目前未使用）
    回傳：整數頁碼或 None（現為 None）
    """
    return None


def articleRow(article0, article, page, body):
    return {
        "article0": article0,
        "article": article,
        "page": page,
        "body": body
    }

def cleanBody(text: str) -> str:
    """
    清理條目 body 的文字格式：
    - 移除空白行
    - 每行左右 trim，並壓縮多餘空白
    - 保留條目之間的單一換行(")\n")，但移除條目內的多餘換行

    目標：移除條文內部的換行符號，每個條文之間只保留一個換行符號。
    """
    # 先分行並去除空行
    rawLines = [l.strip() for l in text.splitlines() if l.strip()]
    # 對每一行壓縮內部空白為單一空白
    normLines = [re.sub(r"\s+", " ", l) for l in rawLines]
    # 用單一換行連接各條目，保留條目之間的換行
    return "\n".join(normLines)

# === 最終清理 body 欄位 ===
def linesToBody(lines):
    # lines: list of (is_item:bool, text:str)
    paragraphs = []
    cur = None
    for is_item, txt in lines:
        # 壓縮內部空白
        txt = re.sub(r"\s+", " ", txt).strip()
        if is_item:
            # item 為獨立段落，先收起前一段
            if cur:
                paragraphs.append(cur)
                cur = None
            paragraphs.append(txt)
        else:
            # 一般段落：合併為 single paragraph
            if cur:
                cur = cur + " " + txt
            else:
                cur = txt
    if cur:
        paragraphs.append(cur)
    # 每個 paragraph 用單一換行分隔
    return "\n".join(paragraphs)

def wordToJson(wordPath, outputPath):
    """
    將單一 .docx 轉成 JSON 並寫入 output_path。

    解析策略（階層）：
    - 第0階：章（只暫存章標題作為補全使用，不輸出章物件）
    - 第1階：條（以 "第...條" 為開頭建立新的輸出物件）
    - 第2階：次標（a,b,c）視為獨立輸出物件，屬於某條之下
    - 條列（如「一、二、三」）視為同一條的條列內容，會以換行保留每項

    每個輸出元素結構：
      {
        "article0": "第X條 完整標題或 a 子標前綴",
        "article": "標題（不含第X條字樣）",
        "page": None,
        "body": "條文內容"
      }
    """
    doc = Document(wordPath)

    # 解析狀態（維持目前的章、條、次標與條列緩衝）
    nowL0Title = None   # 第0階：章標題（僅用於補全條標題）
    nowL1TitleRe = None         # 第1階：目前條物件（dict）
    nowL2TitleRe = None             # 第2階：目前次標物件（dict）
    nowItem = None            # 條列緩衝（記錄最後一個「一、...」項目是否正在續行）

    results = []  # 最終輸出之 list
    prevText = None  # 用於跳過與前一段落完全相同的重複段落

    # loop 讀取 word 檔
    for para in doc.paragraphs:
        paraText = para.text.strip()
        if not paraText:
            # 空段落跳過
            continue

        # ??若與上一段完全相同，視為重複段落跳過（可避免像範例中出現的重複行）
        if prevText is not None and paraText == prevText:
            # 更新 prev_para_text 並跳過處理
            prevText = paraText
            continue

        # === 第0階 title ===
        itemRe = L0TitleGetRe.match(paraText)
        if itemRe:
            resetL2()
            itemRe = itemRe.group(2).strip()
            # 章若標註為 "(刪除)" 則視為無章標題
            nowL0Title = None if itemRe == Deleted else itemRe
            # 新章開始時重置下層狀態
            nowL1TitleRe = None
            nowL2TitleRe = None
            nowItem = None
            prevText = paraText
            continue

        # === 第1階 title ===
        itemRe = L1TitleGetRe.match(paraText)
        if itemRe:
            resetL2()
            text = itemRe.group(2).strip()
            # 條被標註為刪除則忽略該條及其下層
            if text == Deleted:
                nowL1TitleRe = None
                nowL2TitleRe = None
                nowItem = None
                continue

            # debug
            #if title == "聘僱限制":
            #    pdb.set_trace()

            # 若章標存在且條標未包含章名，則在 article 欄位後方加入章名（以括號表示）
            fullTitle = text
            # article 欄位若未包含章名，則在尾端加入 (章名)
            if nowL0Title and nowL0Title not in text:
                fullTitle = f"{text} ({nowL0Title})"

            preText = itemRe.group(1).strip()
            nowL1TitleRe = articleRow(
                f"第{preText}條 {text}",
                fullTitle,
                extractPage(para),
                []
            )
            results.append(nowL1TitleRe)
            # 進入新條後重置次層狀態
            nowL2TitleRe = None
            nowItem = None
            continue

        # === 第1階 item ===
        # 判斷是否為條列（例如「一、」）之起始
        itemRe = L1ItemGetRe.match(paraText)
        # 設定目前要寫入的目標：優先寫入 nowL2，否則寫入 nowL1
        nowTitleRe = nowL2TitleRe if nowL2TitleRe else nowL1TitleRe
        if itemRe:
            resetL2()
            # debug
            pdb.set_trace()

            # 新條列項目：保留該行（包含前綴）
            nowItem = True
            # append as an item-start line
            if nowTitleRe:
                nowTitleRe[Body].append((True, paraText.strip()))
            # 若是子項目，亦同步加入母條的 body（避免重複加入母條標題時會產生多餘項）
            if nowTitleRe is nowL2TitleRe and nowL1TitleRe:
                now_l1_body = nowL1TitleRe[Body]
                now_l1_body.append((True, paraText.strip()))
        else:
            # 非條列開頭，可能是續行或一般段落
            if nowItem:
                # 條列續行：接續到最後一個 item 行
                if nowTitleRe and nowTitleRe[Body]:
                    is_item, prev = nowTitleRe[Body].pop()
                    # prev 必為文字，合併續行
                    merged = prev + " " + paraText.strip()
                    nowTitleRe[Body].append((True, merged))
                else:
                    # 保險 fallback
                    if nowTitleRe:
                        nowTitleRe[Body].append((True, paraText.strip()))
                # 若是子項目，續行也要同步加入母條
                if nowTitleRe is nowL2TitleRe and nowL1TitleRe:
                    is_item, prev = nowL1TitleRe[Body].pop()
                    merged = prev + " " + paraText.strip()
                    nowL1TitleRe[Body].append((True, merged))
            else:
                # 一般段落文字，直接加入 body（非條列項目）
                if nowTitleRe:
                    nowTitleRe[Body].append((False, paraText.strip()))
                    # 若是子項內容，也一併加入母條內容，使母條包含第2階內容
                    if nowTitleRe is nowL2TitleRe and nowL1TitleRe:
                        nowL1TitleRe[Body].append((False, paraText.strip()))

        # === 第2階：次標（a,b,c） ===
        """         
        line = L2Re.match(text)
        if line and now_l1:
            title = line.group(2).strip()
            # 次標若標註為刪除則忽略
            if title == Deleted:
                now_l2 = None
                now_item = None
                continue

            # 第2階主題不需自動加上母條標題
            now_l2 = article_row(
                f"{line.group(1)} {title}",
                title,
                extract_page(para),
                []
            )
            results.append(now_l2)
            # 同步在母條的 body 加入子標題作為段落（保留前綴）
            now_l1[Body].append((True, f"{line.group(1)} {title}"))
            now_item = None
            continue
        """

        # === 第2階 item ===
        # === 偵測括號編號形式的子項（例如 (一)、(二)）視為第2階主題 ===
        itemRe = L2ItemGetRe.match(paraText)
        if itemRe and nowL1TitleRe:
            preText = itemRe.group(1)
            text = itemRe.group(2).strip()
            # 若標註為刪除則忽略
            if text == Deleted:
                nowL2TitleRe = None
                nowItem = None
                continue

            # 檢查父條是否最後一行為 list item（如「四、 請假核准權限」），若是，
            # 則將括號子項視為該 list item 的子內容（不另當作獨立第2階標題），
            # 並建立一個以該 list item 為 article0 的 JSON 物件（若尚未建立）。
            upListItem = None
            if nowL1TitleRe and nowL1TitleRe.get(Body):
                last = nowL1TitleRe[Body][-1]
                if last[0] is True:
                    upListItem = last[1]

            if upListItem:
                # 先找 results 中是否已存在同名的物件（article0 相符）
                existing = None
                for result in results:
                    if result.get(Article0) == upListItem:
                        existing = result
                        break
                if existing is None:
                    # 建立新的第2階物件，以 list item 文本為 article0，article 去除前綴
                    # 去除前綴（例如："四、"）作為 article
                    stripped = re.sub(L1ItemPtnChk, '', upListItem).strip()
                    new_sub = articleRow(
                        upListItem,
                        stripped,
                        extractPage(para),
                        []
                    )
                    results.append(new_sub)
                    nowL2TitleRe = new_sub
                else:
                    nowL2TitleRe = existing

                # 將括號子項的標題也加入父條的 body（作為條列顯示）
                nowL1TitleRe[Body].append((True, f"({preText}) {text}"))
                # 若括號後有內容，視為子項首行內容，加入子項（但不重複加入父條）
                if text:
                    nowL2TitleRe[Body].append((False, text))

                nowItem = None
                prevText = paraText
                continue

            # 若沒有父 list item，則維持原來把括號項目當做第2階主題的行為
            fullTitle = text
            if nowL1TitleRe[Article] not in text:
                fullTitle = f"{nowL1TitleRe[Article]} {text}"

            # 建立第2階物件並加入 results
            nowL2TitleRe = articleRow(
                f"({preText}) {fullTitle}",
                fullTitle,
                extractPage(para),
                []
            )
            results.append(nowL2TitleRe)
            # 同步將該子項目之標題也加入第1階的 body（作為條列顯示）
            nowL1TitleRe[Body].append((True, f"({preText}) {text}"))
            # 若括號後有內容，視為子項首行內容，加入子項（但不重複加入父項）
            if text:
                nowL2TitleRe[Body].append((False, text))

            nowItem = None
            prevText = paraText
            continue

        # === 條文內容處理 ===
        # 條文內容會先放到次標（若存在），否則放到條
        nowTitleRe = nowL2TitleRe if nowL2TitleRe else nowL1TitleRe
        if not nowTitleRe:
            # 無上層標題（孤立文字），忽略
            continue

        # 更新 prev_para_text
        prevText = paraText

    # 寫入 body 欄位
    for result in results:
        lines = result.pop(Body, [])
        result[Body] = linesToBody(lines)

    # === 合併特定父條與其後續括號子項至同一物件 ===
    # 規則：若某物件的 article0 為列表項（如「四、 請假核准權限」），
    # 則緊接在其後且 article0 以括號開頭的物件（如 "(一)..."）視為其子項，
    # 將子項的 body 合併加入父條的 body，並從輸出中移除子項物件。
    merged = []
    i = 0
    # 用來識別 article0 是否為列表項（例如以中文數字開頭並帶「、」）
    # RE_ARTICLE0_LIST = re.compile(L1ItemPtnChk)
    while i < len(results):
        result = results[i]
        # 若為列表項，則合併緊接在後且 article0 以括號開頭的子項
        if L1ItemChkRe.match(result.get(Article0, "")):
            parent = result
            parentBody = (parent.get(Body, "") or "").strip()
            parentParts = [p for p in parentBody.split("\n") if p.strip()]
            i += 1
            # 合併緊接在後且 article0 以括號開頭的物件
            while i < len(results) and results[i].get(Article0, "").strip().startswith("("):
                child = results[i]
                childBody = (child.get(Body, "") or "").strip()
                childHeader = child.get(Article0, "").strip()
                # 準備要加入的段落（header 以及可能的 body）
                if childBody:
                    toAddLines = [childHeader]
                    toAddLines.extend(childBody.split("\n"))
                else:
                    toAddLines = [childHeader]

                # 檢查 parent_parts 是否已包含 child_body 整段（避免重複），或已包含 header
                contains_header = any(p.strip().startswith(childHeader) for p in parentParts)
                contains_body = any(p.strip() == childBody for p in parentParts if childBody)

                if contains_header or contains_body:
                    # 若已有 body 出現但未有 header，嘗試將該行替換為 header+body
                    if contains_body and not contains_header and childBody:
                        new_parts = []
                        replaced = False
                        for p in parentParts:
                            if not replaced and p.strip() == childBody:
                                # 用 header 與 body 替換
                                new_parts.append(childHeader)
                                new_parts.append(childBody)
                                replaced = True
                            else:
                                new_parts.append(p)
                        parentParts = new_parts
                    # 否則跳過加入
                else:
                    # 加到末端
                    parentParts.extend(toAddLines)

                i += 1
            # 移除相鄰重複行
            clean_parts = []
            for p in parentParts:
                ps = p.strip()
                if not ps:
                    continue
                if clean_parts and clean_parts[-1].strip() == ps:
                    continue
                clean_parts.append(ps)
            parent[Body] = "\n".join(clean_parts)
            merged.append(parent)
        else:
            merged.append(result)
            i += 1

    results = merged

    # 進一步清理：移除父條 body 中因合併而產生的重複段落
    # PAREN_HEADER_RE = re.compile(L2ItemPtnChk)
    for result in results:
        if not result.get(Body):
            continue
        parts = result[Body].split("\n")
        cleaned = []
        for idx, itemRe in enumerate(parts):
            sline = itemRe.strip()
            if not sline:
                continue
            # 若下一行存在且為帶括號前綴，且下一行剝掉前綴後等於本行，則跳過本行
            if idx + 1 < len(parts):
                next_line = parts[idx + 1].strip()
                itemRe = L2ItemChkRe.match(next_line)
                if itemRe:
                    next_without_header = L2ItemChkRe.sub('', next_line).strip()
                    if next_without_header and next_without_header == sline:
                        # skip current line to avoid duplicated content before '(一) ...'
                        continue
            # 若與前一已加入的行相同，跳過
            if cleaned and cleaned[-1].strip() == sline:
                continue
            cleaned.append(sline)
        result[Body] = "\n".join(cleaned)

    # 修正：確保父條中包含的子項若對應到某個以括號為 article0 的物件，
    # 則在對應行前保留該子項的前置文字（如「(五)」）。
    # 建立從子項內容到其前置文字的對照表
    # PAREN_LINE_RE = re.compile(L2ItemPtnGet)
    paren_map = {}
    # 掃描所有結果的 article0 / body 中可能的括號前綴與其後內容
    for result in results:
        # article0 若以括號開頭
        a0 = result.get(Article0, "").strip()
        if a0.startswith("("):
            itemRe = L2ItemGetRe.match(a0)
            if itemRe:
                header = f"({itemRe.group(1)})"
                text = itemRe.group(2).strip()
                if text:
                    paren_map[text] = header
                paren_map[header] = header
        # body 中每一行
        body = result.get(Body, "") or ""
        for itemRe in body.split("\n"):
            lm = L2ItemGetRe.match(itemRe.strip())
            if lm:
                header = f"({lm.group(1)})"
                text = lm.group(2).strip()
                if text:
                    paren_map[text] = header
                paren_map[header] = header

    # 以 paren_map 為基礎，在父條中補回缺失的括號前綴（泛用，不硬編碼）
    if paren_map:
        for result in results:
            body = result.get(Body, "")
            if not body:
                continue
            lines = body.split("\n")
            new_lines = []
            for ln in lines:
                s = ln.strip()
                added = False
                # 對照表中較長字串優先匹配，避免短字串誤命中
                for key in sorted(paren_map.keys(), key=lambda k: -len(k)):
                    if not key:
                        continue
                    if key in s:
                        header = paren_map[key]
                        # 若該行尚未以 header 開頭，插入 header 行
                        if not s.startswith(header):
                            if not (new_lines and new_lines[-1].strip() == header):
                                new_lines.append(header)
                        new_lines.append(s)
                        added = True
                        break
                if not added:
                    new_lines.append(s)
            # 去重相鄰重複
            cleaned = []
            for ln in new_lines:
                if cleaned and cleaned[-1].strip() == ln.strip():
                    continue
                cleaned.append(ln)
            result[Body] = "\n".join(cleaned)

    # 合併孤立的括號前置標題行（例如單獨一行的 "(五)"）與其下一行
    # PAREN_HEADER_ONLY_RE = re.compile(L2ItemPtnChk)
    for result in results:
        body = result.get(Body, "")
        if not body:
            continue
        lines = body.split("\n")
        new_lines = []
        i = 0
        while i < len(lines):
            itemRe = lines[i].strip()
            if L2ItemChkRe.match(itemRe) and i + 1 < len(lines):
                # 合併此行與下一行
                next_line = lines[i + 1].strip()
                merged = itemRe + " " + next_line if next_line else itemRe
                new_lines.append(merged)
                i += 2
            else:
                new_lines.append(itemRe)
                i += 1
        result[Body] = "\n".join(new_lines)

    # 最終：壓縮多重換行並 trim
    for result in results:
        body = result.get(Body, "")
        if body:
            # 將兩個或以上的換行縮為一個，並去除首尾空白
            result[Body] = re.sub(r"\n{2,}", "\n", body).strip()

    # 寫入 JSON檔，使用 ensure_ascii=False 以保留中文
    with open(outputPath, "w", encoding="utf-8") as file:
        json.dump(results, file, ensure_ascii=False, indent=2)


# === 主程式 ===
if __name__ == "__main__":
    # 範例執行：開發時可改為從命令列參數讀入路徑
    inputPath="input/工作規則-easy.docx",
    outputDir="output"

    #if not inputPath.lower().endswith(".docx"):
    #    raise ValueError("只接受 .docx 檔案")

    if not os.path.isfile(inputPath):
        raise FileNotFoundError(f"找不到檔案：{inputPath}")

    os.makedirs(outputDir, exist_ok=True)

    fname = os.path.basename(inputPath)
    output_path = os.path.join(
        outputDir,
        fname.replace(".docx", ".json")
    )

    wordToJson(inputPath, output_path)
    print(f"Converted: {fname}")
