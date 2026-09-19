import pandas as pd
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

EXCEL_FILE = "以撒的结合忏悔+_全道具信息表.xlsx"
STATE_FILE = "rating_state.json"
SHEET_NAME = "道具信息"

def load_items(excel_path):
    """读取 Excel，返回道具列表（已排除里该隐评级为 -1 的道具）"""
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    required_cols = ['道具ID', '英文名', '中文名', '道具等级（品质）', '道具介绍', '主动/被动道具', '里该隐评级']
    for col in required_cols:
        if col not in df.columns:
            print(f"错误：Excel 中找不到列 '{col}'，请检查表头。")
            sys.exit(1)
    df = df[df['里该隐评级'] != -1].copy()
    df = df.reset_index(drop=True)
    items = []
    for _, row in df.iterrows():
        item = {
            'id': int(row['道具ID']),
            '英文名': str(row['英文名']),
            '中文名': str(row['中文名']),
            '品质': int(row['道具等级（品质）']),
            '介绍': str(row['道具介绍']),
            '类型': str(row['主动/被动道具']),
            '里该隐评级': row['里该隐评级']
        }
        items.append(item)
    return items

def save_state(state):
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def format_item(item):
    lines = [
        f"ID: {item['id']}",
        f"中文名: {item['中文名']}",
        f"英文名: {item['英文名']}",
        f"品质: {item['品质']}",
        f"类型: {item['类型']}",
        f"里该隐评级: {item['里该隐评级']}",
        f"介绍: {item['介绍']}"
    ]
    return "\n".join(lines)

def main():
    print("=== 里该隐道具评级（两两比较，从低到高，选择更好的） ===")
    if not os.path.exists(EXCEL_FILE):
        print(f"找不到文件：{EXCEL_FILE}")
        excel_path = input("请输入 Excel 文件路径：").strip()
        if not os.path.exists(excel_path):
            print("文件不存在，退出。")
            return
    else:
        excel_path = EXCEL_FILE

    print("正在加载道具数据...")
    items = load_items(excel_path)
    item_dict = {item['id']: item for item in items}
    total = len(items)
    print(f"共加载 {total} 个可合成道具（已排除里该隐评级为 -1 的道具）。")

    state = load_state()
    if state:
        print("发现保存的状态。")
        choice = input("是否继续上次的进度？(y/n，默认 y)：").strip().lower()
        if choice == 'n':
            state = None
            if os.path.exists(STATE_FILE):
                os.remove(STATE_FILE)
            print("已删除旧状态，重新开始。")
        else:
            print("继续上次进度。")

    if state is None:
        state = {
            'sorted_ids': [],
            'next_index': 0,
            'current_item_id': None,
            'low': 0,
            'high': 0
        }
        if total > 0:
            state['sorted_ids'].append(items[0]['id'])
            state['next_index'] = 1
        save_state(state)

    while True:
        sorted_ids = state['sorted_ids']
        next_index = state['next_index']
        current_item_id = state['current_item_id']

        if current_item_id is None:
            if next_index >= total:
                print("所有道具已评级完成！")
                break
            current_item = items[next_index]
            current_item_id = current_item['id']
            if not sorted_ids:
                sorted_ids.append(current_item_id)
                state['next_index'] = next_index + 1
                state['current_item_id'] = None
                save_state(state)
                continue
            state['current_item_id'] = current_item_id
            state['low'] = 0
            state['high'] = len(sorted_ids) - 1
            save_state(state)
            continue

        low = state['low']
        high = state['high']

        if low > high:
            pos = low
            sorted_ids.insert(pos, current_item_id)
            state['next_index'] = next_index + 1
            state['current_item_id'] = None
            state['low'] = 0
            state['high'] = 0
            save_state(state)
            item = item_dict[current_item_id]
            print(f"\n>>> 已插入：{item['中文名']}（ID: {item['id']}）到位置 {pos+1}")
            print(f"当前进度：{len(sorted_ids)} / {total}")
            continue

        mid = (low + high) // 2
        mid_id = sorted_ids[mid]
        current_item = item_dict[current_item_id]
        mid_item = item_dict[mid_id]

        print("\n" + "=" * 60)
        print("待插入道具：")
        print(format_item(current_item))
        print("\n" + "-" * 60)
        print("已排序道具（中间位置）：")
        print(format_item(mid_item))
        print("=" * 60)
        print("哪个道具的里该隐评级更好（更高）？")
        print("输入 1：待插入道具更好")
        print("输入 2：已排序道具更好")
        print("输入 =：两者相等（按 ID 升序）")
        print("输入 q：保存并退出")
        choice = input("你的选择：").strip()

        if choice == 'q':
            save_state(state)
            print("已保存，退出。")
            return
        elif choice == '1':
            # 待插入更好 → 应放在右边
            state['low'] = mid + 1
        elif choice == '2':
            # 已排序道具更好 → 待插入更差 → 应放在左边
            state['high'] = mid - 1
        elif choice == '=':
            if current_item_id < mid_id:
                state['high'] = mid - 1
            else:
                state['low'] = mid + 1
        else:
            print("无效输入，请重新输入。")
            continue

        save_state(state)

    print("\n所有道具评级完成！")
    result_data = []
    for rank, item_id in enumerate(state['sorted_ids'], start=1):
        item = item_dict[item_id]
        result_data.append({
            '排名（从低到高）': rank,
            '道具ID': item['id'],
            '中文名': item['中文名'],
            '英文名': item['英文名'],
            '品质': item['品质'],
            '类型': item['类型'],
            '里该隐评级': item['里该隐评级'],
            '介绍': item['介绍']
        })
    result_df = pd.DataFrame(result_data)
    output_file = "里该隐评级_从低到高.xlsx"
    result_df.to_excel(output_file, index=False)
    print(f"结果已保存至：{output_file}")
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        print("状态文件已删除。")

if __name__ == '__main__':
    main()