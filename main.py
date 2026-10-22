"""
main.py - 命令列入口點
使用方式：
    python main.py <CAD檔案路徑>
    例如：python main.py sample.step
"""

import sys
import json
from cad_parser import parse_cad_bounding_box


def main():
    if len(sys.argv) < 2:
        print("【使用說明】")
        print("  請輸入要解析的 CAD 檔案路徑：")
        print("  python main.py <file_path>")
        print("\n範例：")
        print("  python main.py ./demo.step")
        sys.exit(1)

    target_file = sys.argv[1]
    print(f"正在解析檔案: {target_file} ...")

    result = parse_cad_bounding_box(target_file)

    if result.get("status") == "success":
        print("\n================ 解析成功 ================")
        print(f"檔案名稱  : {result['file_name']}")
        print(f"原始尺寸  : X={result['raw_dimensions']['x_axis_mm']} mm, "
              f"Y={result['raw_dimensions']['y_axis_mm']} mm, "
              f"Z={result['raw_dimensions']['z_axis_mm']} mm")
        print("------------------------------------------")
        stock = result['stock_dimensions']
        print(f"建議胚料尺寸 (長x寬x高): {stock['length_mm']} x {stock['width_mm']} x {stock['height_mm']} mm")
        print("==========================================\n")

        print("JSON 格式輸出：")
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"\n[錯誤] 解析失敗：{result.get('error_message')}")


if __name__ == "__main__":
    main()