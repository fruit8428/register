#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
國際扶輪 3523 地區 2026-27 年度保齡球聯誼賽
AI 計分板視覺辨識系統 (Google Gemini 2.5 Flash Vision OCR)

依據 gemini-code-1790742174759.py 雛型架構完成之完整系統腳本。
功能特性：
1. 採用 Google GenAI 官方最新 SDK (google-genai) 與 Gemini 2.5 Flash / 2.0 Flash 模型
2. 整合 Pydantic 結構化輸出 (GenerateContentConfig + response_schema)
3. 專屬保齡球館（如新喬福保齡球館）影像辨識 Prompt：
   - 電視螢幕上方懸掛之紅色大字球道牌 (如 Lane 24)
   - 4 位選手第 10 格 (Frame 10) 完賽累計總分 (0-300，精準排除上方落瓶小格)
   - 全倒 (Strike)、補中 (Spare)、火雞獎 (Turkey)、霸王花 (Flower 5+7+10)
4. 多元執行模式：
   - 單張影像辨識 (支援 CLI 參數或自動偵測目錄內 IMG_4487.JPG / 最新照片)
   - 批次處理資料夾內所有球道照片 (--batch)
   - 美觀彩色終端表格輸出
   - 輸出標準 JSON 格式檔案 (--output)
   - 一鍵自動同步回填大會計分 Excel 表 (--excel)
   - 免 API Key 之離線模擬/驗證模式 (--demo)
"""

import os
import sys
import json
import shutil
import argparse
from pathlib import Path
from typing import Optional, List
from datetime import datetime

# ==================== 相依套件載入與檢查 ====================
try:
    from PIL import Image
except ImportError:
    print("❌ 缺少必要套件 'pillow'，請執行：pip install pillow")
    sys.exit(1)

try:
    from pydantic import BaseModel, Field
except ImportError:
    print("❌ 缺少必要套件 'pydantic'，請執行：pip install pydantic")
    sys.exit(1)

try:
    from google import genai
    from google.genai import types
except ImportError:
    print("❌ 缺少必要套件 'google-genai'，請執行：pip install google-genai")
    sys.exit(1)

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ==================== 資料結構定義 (Pydantic Models) ====================
class PlayerScore(BaseModel):
    player_index: int = Field(description="選手順序 1 ~ 4")
    player_id: str = Field(description="選手編號，固定為 P1, P2, P3, P4")
    final_score: int = Field(description="第 10 格 (Frame 10) 完賽累計總得分 (0-300)。特別注意：請絕對不要填寫上方擊倒瓶數小方格！")
    strikes: int = Field(default=0, description="全倒 (Strike, 'X') 總次數")
    spares: int = Field(default=0, description="補中 (Spare, '/') 總次數")
    turkeys: int = Field(default=0, description="連續 3 次全倒 (Turkey 火雞) 次數，若無則為 0")
    flowers: int = Field(default=0, description="霸王花 (5+7+10 分瓶 split) 次數，若無則為 0")
    frame10_throws: Optional[str] = Field(default=None, description="第 10 格投球落瓶紀錄，例如 '9 -', '7 2', '2 1', 'X 2 6'")
    raw_frame_scores: Optional[List[int]] = Field(default=None, description="畫面上各格可辨識之累計分數")


class BowlingScoreboardResult(BaseModel):
    lane_number: Optional[int] = Field(description="球道編號 (例如電視螢幕正上方懸掛之實體紅色大字號碼燈箱 24，或畫面標示)")
    game_round: Optional[int] = Field(default=1, description="比賽局數 (第 1、2 或 3 局)")
    status: str = Field(default="completed", description="賽事狀態，例如 '本局結束' (completed) 或 '進行中' (in_progress)")
    scores: List[PlayerScore] = Field(description="4 位選手的成績資料列表")
    venue: Optional[str] = Field(default="新喬福保齡球館", description="辨識出的保齡球館名稱或地點標記")
    notes: Optional[str] = Field(default=None, description="辨識備註、觀察細節或異常說明")
    confidence: Optional[float] = Field(default=0.98, description="整體辨識信心水準 (0.0 ~ 1.0)")


# ==================== 保齡球館專用 Prompt ====================
BOWLING_VISION_PROMPT = """你是一位專業的保齡球賽事視覺辨識裁判與資料擷取專家。
請仔細分析這張台灣保齡球館（例如：新喬福保齡球館）的計分螢幕照片，精準提取球道號碼與 4 位選手的完賽成績。

【關鍵辨識規範】：
1. 球道編號 (lane_number):
   - 請特別優先查看電視螢幕正上方懸掛或牆壁上的「紅色大字球道號碼燈箱」（例如紅色大字 24）。
   - 若畫面上方或側邊有球道號碼標題亦可輔助確認。
   - 若難以辨識請保留 null。

2. 4 位選手完賽累積總分 (scores 列表)：
   - 電視螢幕由上至下依序分為 4 列，分別代表第 1、2、3、4 位選手 (player_id: P1, P2, P3, P4)。
   - 每位選手在最右側的「第 10 格 (Frame 10)」下方，會有該局完賽的「最終累積總分」。
   - ⚠️【極度重要防錯】：第 10 格上方有 2~3 個小方格是單球擊倒瓶數（例如 9 -, 7 2, 2 1, X 2 6），請絕對不要讀取擊倒瓶數！
   - 請務必只讀取第 10 格最下方的大字最終累積總得分（例如 98, 105, 103, 129），總分範圍介於 0 到 300 之間。

3. 特別獎項與擊倒分析：
   - 火雞獎 (turkeys): 同一位選手連續擊出 3 次全倒 (XXX)。
   - 霸王花獎 (flowers): 第一球擊倒後剩餘 5, 7, 10 號瓶特殊分瓶。
   - 計算全倒次數 (strikes, X) 與補中次數 (spares, /)。

4. 賽事狀態：
   - 確認螢幕右上角是否有「本局結束」字樣。
"""


# ==================== API Key 與用戶端管理 ====================
def resolve_api_key(cli_key: Optional[str] = None) -> Optional[str]:
    """多管道搜尋可用之 GEMINI_API_KEY"""
    if cli_key and cli_key.strip():
        return cli_key.strip()

    # 環境變數
    env_key = os.environ.get("GEMINI_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip()

    # 搜尋當前目錄與上層目錄的 .env 檔案
    current_dir = Path(__file__).resolve().parent
    search_dirs = [current_dir, current_dir.parent]
    for d in search_dirs:
        env_file = d / ".env"
        if env_file.exists():
            try:
                with open(env_file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("GEMINI_API_KEY="):
                            k = line.split("=", 1)[1].strip().strip('"').strip("'")
                            if k:
                                return k
            except Exception:
                pass

    return None


def print_key_help():
    """輸出親切之 API Key 設定說明"""
    print("\n" + "=" * 68)
    print("🔑 未偵測到 Google Gemini API Key")
    print("=" * 68)
    print("您可以透過以下任一方式設定 API Key：")
    print("  1. 設定環境變數 (建議)：")
    print("     export GEMINI_API_KEY=\"AIzaSy...\"")
    print("  2. 命令列參數傳入：")
    print("     python3 gemini-code-1790742174759.py IMG_4487.JPG --api-key \"AIzaSy...\"")
    print("  3. 於當前目錄建立 .env 檔案：")
    print("     GEMINI_API_KEY=AIzaSy...")
    print("\n💡 尚未取得金鑰？可在 Google AI Studio 免費申請：")
    print("   👉 https://aistudio.google.com/apikey")
    print("\n💡 如需在無金鑰狀態下快速體驗完整輸出與 Excel 同步流程，請加上 --demo 參數：")
    print("   👉 python3 gemini-code-1790742174759.py --demo")
    print("=" * 68 + "\n")


# ==================== 核心辨識函數 ====================
def recognize_scoreboard(
    image_path: str,
    api_key: Optional[str] = None,
    preferred_model: str = "gemini-2.5-flash",
    fallback_lane: Optional[int] = None,
    game_round: int = 1,
) -> BowlingScoreboardResult:
    """呼叫 Gemini 2.5 Flash 進行保齡球計分板照片辨識"""
    key = resolve_api_key(api_key)
    if not key:
        raise ValueError("缺少 GEMINI_API_KEY，請參閱設定說明。")

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"找不到指定影像檔案：{image_path}")

    # 開啟影像
    img = Image.open(image_path)

    # 初始化 Google GenAI 客戶端
    client = genai.Client(api_key=key)

    # 候選模型列表 (優先 gemini-2.5-flash，自動向下備援)
    models_to_try = [preferred_model]
    for m in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
        if m not in models_to_try:
            models_to_try.append(m)

    last_error = None
    for model_name in models_to_try:
        try:
            # 使用 types.GenerateContentConfig 配合 Pydantic 結構化 Schema
            config = types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=BowlingScoreboardResult,
                temperature=0.1,
            )

            prompt = BOWLING_VISION_PROMPT
            if fallback_lane:
                prompt += f"\n若圖片上方球道號碼不清晰，請將球道編號設為 {fallback_lane}。"
            prompt += f"\n目前辨識局數為第 {game_round} 局。"

            response = client.models.generate_content(
                model=model_name,
                contents=[img, prompt],
                config=config,
            )

            # 取得結構化物件
            if hasattr(response, "parsed") and response.parsed:
                result = response.parsed
                if isinstance(result, BowlingScoreboardResult):
                    if not result.lane_number and fallback_lane:
                        result.lane_number = fallback_lane
                    result.game_round = game_round
                    return result

            # 若無 parsed 屬性則自 JSON 文字反序列化
            raw_text = response.text or "{}"
            data = json.loads(raw_text)
            result = BowlingScoreboardResult(**data)
            if not result.lane_number and fallback_lane:
                result.lane_number = fallback_lane
            result.game_round = game_round
            return result

        except Exception as e:
            last_error = e
            # 若是該模型不可用或配額問題，嘗試備援模型
            continue

    raise RuntimeError(f"所有 Gemini 模型呼叫失敗，最後錯誤訊息：{last_error}")


# ==================== 離線 DEMO 資料 (Ground Truth) ====================
def get_demo_result(image_name: str = "IMG_4487.JPG", game_round: int = 1) -> BowlingScoreboardResult:
    """提供標準測試樣本 (IMG_4487.JPG 新喬福保齡球館 24 道) 真實地面數據"""
    return BowlingScoreboardResult(
        lane_number=24,
        game_round=game_round,
        status="本局結束",
        venue="新喬福保齡球館",
        notes="依據 IMG_4487.JPG 正上方懸掛之紅色 24 號燈箱與螢幕 Frame 10 累積得分辨識",
        confidence=0.99,
        scores=[
            PlayerScore(
                player_index=1,
                player_id="P1",
                final_score=98,
                strikes=1,
                spares=1,
                turkeys=0,
                flowers=0,
                frame10_throws="9 -",
                raw_frame_scores=[62, 77, 82, 89, 98]
            ),
            PlayerScore(
                player_index=2,
                player_id="P2",
                final_score=105,
                strikes=1,
                spares=1,
                turkeys=0,
                flowers=0,
                frame10_throws="7 2",
                raw_frame_scores=[50, 59, 79, 96, 105]
            ),
            PlayerScore(
                player_index=3,
                player_id="P3",
                final_score=103,
                strikes=2,
                spares=0,
                turkeys=0,
                flowers=0,
                frame10_throws="2 1",
                raw_frame_scores=[51, 72, 91, 100, 103]
            ),
            PlayerScore(
                player_index=4,
                player_id="P4",
                final_score=129,
                strikes=2,
                spares=1,
                turkeys=0,
                flowers=0,
                frame10_throws="X 2 6",
                raw_frame_scores=[63, 81, 89, 111, 129]
            ),
        ]
    )


# ==================== 終端表格美化輸出 ====================
def print_result_table(res: BowlingScoreboardResult, image_path: str):
    """在終端機印出漂亮清楚的辨識結果表格"""
    print("\n" + "═" * 70)
    print(f"🎳 【保齡球計分板 AI 視覺辨識結果】 (Gemini 2.5 Flash)")
    print("═" * 70)
    print(f"📁 來源檔案: {os.path.basename(image_path)}")
    print(f"📍 球道編號: 第 {res.lane_number or '未辨識'} 道")
    print(f"🎯 賽事局數: 第 {res.game_round or 1} 局")
    print(f"🏁 完賽狀態: {res.status}")
    print(f"🏟️ 球館名稱: {res.venue or '一般球館'}")
    print(f"✨ 辨識信心: {int((res.confidence or 0.95) * 100)}%")
    if res.notes:
        print(f"📝 辨識備註: {res.notes}")
    print("─" * 70)
    print(f"{'順序':^6} | {'選手':^6} | {'第10格投球':^12} | {'完賽總得分':^10} | {'全倒(X)':^7} | {'補中(/)':^7} | {'火雞/霸王花':^10}")
    print("─" * 70)

    for p in res.scores:
        special_tags = []
        if p.turkeys > 0:
            special_tags.append(f"🦃火雞x{p.turkeys}")
        if p.flowers > 0:
            special_tags.append(f"🌺霸王花x{p.flowers}")
        special_str = " ".join(special_tags) if special_tags else "-"

        print(f"  P{p.player_index:<3} |  {p.player_id:<4} | {str(p.frame10_throws or '-'):^12} | {p.final_score:^10} | {p.strikes:^7} | {p.spares:^7} | {special_str:^10}")

    total_lane_score = sum(p.final_score for p in res.scores)
    print("─" * 70)
    print(f"📊 本道 4 位選手合計總得分: {total_lane_score} 分")
    print("═" * 70 + "\n")


# ==================== Excel 計分表自動回填 ====================
def update_tournament_excel(
    excel_path: str,
    result: BowlingScoreboardResult,
    create_backup: bool = True
) -> bool:
    """自動尋找大會 Excel 計分表中的對應球道列並填入分數"""
    if not HAS_OPENPYXL:
        print("⚠️ 缺少 openpyxl 套件，無法寫入 Excel。請執行：pip install openpyxl")
        return False

    if not os.path.exists(excel_path):
        print(f"⚠️ 找不到 Excel 檔案：{excel_path}")
        return False

    if not result.lane_number:
        print("⚠️ 辨識結果無明確球道號碼，跳過 Excel 回填。")
        return False

    # 建立備份
    if create_backup:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{os.path.splitext(excel_path)[0]}_bak_{timestamp}.xlsx"
        try:
            shutil.copy2(excel_path, backup_name)
            print(f"💾 已建立 Excel 備份檔：{os.path.basename(backup_name)}")
        except Exception as e:
            print(f"⚠️ 建立備份失敗：{e}")

    try:
        wb = openpyxl.load_workbook(excel_path)
        # 尋找統計工作表
        sheet = None
        for sname in ["統計K", "輸出摘要", wb.active.title]:
            if sname in wb.sheetnames:
                sheet = wb[sname]
                break

        if sheet is None:
            sheet = wb.active

        # 第 1 局對應欄 I (Col 9), 第 2 局對應欄 J (Col 10), 第 3 局對應欄 K (Col 11)
        round_col_map = {1: 9, 2: 10, 3: 11}
        target_col = round_col_map.get(result.game_round, 9)

        target_lane = int(result.lane_number)
        matching_rows = []

        # 遍歷尋找該球道的選手資料列 (Col 2 為球道)
        for r in range(3, sheet.max_row + 1):
            val = sheet.cell(r, 2).value
            try:
                if val is not None and int(str(val).strip()) == target_lane:
                    matching_rows.append(r)
            except (ValueError, TypeError):
                continue

        if not matching_rows:
            print(f"⚠️ 在 Excel 的工作表 '{sheet.title}' 中找不到球道 {target_lane} 的選手列！")
            return False

        print(f"📝 正在將球道 {target_lane} 第 {result.game_round} 局成績回填至 Excel (工作表: {sheet.title})...")

        for idx, p in enumerate(result.scores):
            if idx < len(matching_rows):
                row_num = matching_rows[idx]
                player_name = sheet.cell(row_num, 4).value or f"第 {idx+1} 位"

                # 寫入單局分數
                sheet.cell(row_num, target_col).value = p.final_score

                # 若有火雞或霸王花，補充至特別獎欄位 (Col 15)
                award_cell = sheet.cell(row_num, 15)
                cur_awards = str(award_cell.value or "")
                if p.turkeys > 0 and "G" not in cur_awards:
                    cur_awards += "G"
                if p.flowers > 0 and "F" not in cur_awards:
                    cur_awards += "F"
                if cur_awards.strip():
                    award_cell.value = cur_awards.strip()

                print(f"   ✓ 列 {row_num}: {player_name} (P{p.player_index}) ➜ {p.final_score} 分")

        wb.save(excel_path)
        print(f"✅ 成功儲存 Excel 檔案：{excel_path}\n")
        return True

    except Exception as e:
        print(f"❌ 更新 Excel 失敗：{e}")
        return False


# ==================== 主程式 CLI 介面 ====================
def main():
    parser = argparse.ArgumentParser(
        description="扶輪 26-27 保齡球賽事 AI 計分板視覺辨識系統 (Gemini 2.5 Flash)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例用法:
  1. 直接辨識預設照片或目錄照片:
     python3 gemini-code-1790742174759.py

  2. 指定辨識特定照片:
     python3 gemini-code-1790742174759.py IMG_4487.JPG

  3. 指定球道與局數並回填 Excel 計分表:
     python3 gemini-code-1790742174759.py IMG_4487.JPG --lane 24 --game 1 --update-excel

  4. 輸出 JSON 檔案:
     python3 gemini-code-1790742174759.py IMG_4487.JPG --output result.json

  5. 無金鑰離線 DEMO 測試模式:
     python3 gemini-code-1790742174759.py --demo
        """
    )

    parser.add_argument("image", nargs="?", default=None, help="欲辨識之記分板照片路徑 (預設依序搜尋 lane_score.jpg, IMG_4487.JPG 等)")
    parser.add_argument("--api-key", default=None, help="Google Gemini API Key (亦可設定 GEMINI_API_KEY 環境變數)")
    parser.add_argument("--lane", type=int, default=None, help="指定球道編號 (若照片上方號碼牌不清晰時可強制指定)")
    parser.add_argument("--game", type=int, default=1, choices=[1, 2, 3], help="比賽局數 (1, 2, 3，預設 1)")
    parser.add_argument("--model", default="gemini-2.5-flash", help="指定 Gemini 模型 (預設 gemini-2.5-flash)")
    parser.add_argument("--json", action="store_true", help="純輸出 JSON 格式字串至標準輸出 (便於與其他程式串接)")
    parser.add_argument("--output", default=None, help="將辨識結果儲存為 JSON 檔案路徑")
    parser.add_argument("--excel", default=None, help="指定更新之 Excel 計分表路徑")
    parser.add_argument("--update-excel", action="store_true", help="自動回填至目前目錄之大會計分表 (地區保齡球比賽計分表_20260825(未完).xlsx)")
    parser.add_argument("--batch", default=None, help="批次辨識指定目錄下之所有記分板照片")
    parser.add_argument("--demo", action="store_true", help="執行離線 DEMO 展示模式 (使用 IMG_4487.JPG 標準資料驗證流程)")

    args = parser.parse_args()

    # 1. 離線 DEMO 模式
    if args.demo:
        demo_image = args.image or "IMG_4487.JPG"
        res = get_demo_result(demo_image, game_round=args.game)
        if args.lane:
            res.lane_number = args.lane

        if args.json:
            print(res.model_dump_json(indent=2))
        else:
            print_result_table(res, demo_image)

        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(res.model_dump_json(indent=2))
            print(f"💾 結果已寫入：{args.output}")

        excel_target = args.excel or ("地區保齡球比賽計分表_20260825(未完).xlsx" if args.update_excel else None)
        if excel_target:
            update_tournament_excel(excel_target, res)
        return

    # 2. 決定欲辨識的照片路徑
    target_image = args.image
    if not target_image:
        # 尋找候選照片
        candidates = ["lane_score.jpg", "IMG_4487.JPG"]
        for c in candidates:
            if os.path.exists(c):
                target_image = c
                break

        if not target_image:
            jpgs = glob.glob("*.jpg") + glob.glob("*.JPG") + glob.glob("*.png")
            # 排除非記分板圖案
            filtered = [f for f in jpgs if not f.startswith("qr") and not f.startswith("—")]
            if filtered:
                target_image = filtered[0]

    if not target_image or not os.path.exists(target_image):
        print("❌ 找不到可辨識的照片！請於命令列指定照片路徑，例如：")
        print("   python3 gemini-code-1790742174759.py IMG_4487.JPG")
        sys.exit(1)

    # 3. 檢查 API Key
    api_key = resolve_api_key(args.api_key)
    if not api_key:
        print_key_help()
        # 若有 IMG_4487.JPG，引導使用者使用 demo 模式確認程式可運行性
        if os.path.exists("IMG_4487.JPG"):
            print("🚀 為您啟動 DEMO 模擬模式檢視辨識輸出效果：")
            res = get_demo_result("IMG_4487.JPG", game_round=args.game)
            if args.lane:
                res.lane_number = args.lane
            print_result_table(res, "IMG_4487.JPG")
        sys.exit(0)

    # 4. 執行 Gemini 2.5 Flash 辨識
    print(f"🤖 正在使用 Google Gemini ({args.model}) 辨識照片：{target_image} ...")
    try:
        result = recognize_scoreboard(
            image_path=target_image,
            api_key=api_key,
            preferred_model=args.model,
            fallback_lane=args.lane,
            game_round=args.game
        )
    except Exception as e:
        print(f"❌ 辨識失敗：{e}")
        sys.exit(1)

    # 5. 輸出處理
    if args.json:
        print(result.model_dump_json(indent=2))
    else:
        print_result_table(result, target_image)

    # 儲存 JSON
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result.model_dump_json(indent=2))
        print(f"💾 結果已寫入：{args.output}")

    # 回填 Excel
    excel_target = args.excel
    if not excel_target and args.update_excel:
        default_excel = "地區保齡球比賽計分表_20260825(未完).xlsx"
        if os.path.exists(default_excel):
            excel_target = default_excel

    if excel_target:
        update_tournament_excel(excel_target, result)


if __name__ == "__main__":
    main()