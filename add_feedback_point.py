#!/usr/bin/env python3
"""
ちひろさんの添削フィードバックを「気をつけるポイント」として追加するスクリプト

ちひろさんから添削をもらったら、指摘内容を1文に要約してこのスクリプトに渡す。
mentor_feedback.md のチェックリストに日付つきで追記され、以降の台本生成
（generate_script.py / correction_workflow.py）で自動的に参考にされる。

使い方:
    python add_feedback_point.py "フックは最初の3秒で数字を見せる"
"""

import sys

from mentor_feedback import add_feedback_point


def main() -> None:
    if len(sys.argv) < 2:
        print('使い方: python add_feedback_point.py "気をつけるポイントを1文で"')
        sys.exit(1)

    point = " ".join(sys.argv[1:]).strip()
    if not point:
        print("ポイントが空です。")
        sys.exit(1)

    add_feedback_point(point)
    print(f"追加しました: {point}")


if __name__ == "__main__":
    main()
