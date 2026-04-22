"""Generate April narration Word document."""
from pathlib import Path

from narration_04 import NARRATION_04
from narration_doc import build_narration_doc

SLIDE_TITLES = [
    "表紙",
    "本日の目的と到達目標",
    "1. そもそも「介護倫理」とは",
    "2. 倫理綱領 ① 利用者本位・専門性・情報管理",
    "3. 倫理綱領 ② 連携協働・代弁と育成",
    "4. 利用者の尊厳を守るとは（○×比較）",
    "5. 事例① 呼び方の問題",
    "6. 事例② 入浴介助の場面",
    "7. 守秘義務の基礎（法的根拠）",
    "8. 守秘義務の対象・範囲",
    "9. もし守秘義務違反が起きたら",
    "10. 事例③ 居酒屋での会話",
    "11. 事例④ SNSへの写真投稿",
    "12. 事例⑤ 家族からの電話照会",
    "13. 倫理的ジレンマに出会ったら（判断の手順）",
    "14. 事例⑥ 本人の希望 vs 安全",
    "15. 関係法令の最低ライン",
    "自己チェック",
    "まとめ・今日から実践すること",
    "質疑応答",
]


def main():
    out = (
        Path(__file__).resolve().parent.parent
        / "法定研修" / "04_倫理・法令遵守_ナレーション原稿.docx"
    )
    build_narration_doc(
        out,
        facility="みんなの介護ゆうあい",
        month="4月",
        theme="倫理・法令遵守",
        subtitle="利用者の尊厳を守り、守秘義務を徹底する",
        slides=list(zip(SLIDE_TITLES, NARRATION_04)),
    )
    print(f"created: {out}")


if __name__ == "__main__":
    main()
