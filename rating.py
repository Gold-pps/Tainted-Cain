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
OUTPUT_FILE = "里该隐评级_从低到高.xlsx"
PROGRESS_FILE = "里该隐评级_进度.xlsx"
SHEET_NAME = "道具信息"

HELP_TIP = (
    "比较过程中随时可用：\n"
    "  l [起始名次] [数量]  查看已排序清单（带中文名，省略参数 = 全部）\n"
    "  f <ID或名字>         查某个道具排在第几名\n"
    "  r  重新评级某个道具（选错了用它改，会把它摘出来重新比较）\n"
    "  i  手动指定名次插入（自由插入，不比较）\n"
    "  q  保存并退出\n"
    "不启动比较、只想看进度：\n"
    "  uv run rating.py --show [起始名次] [数量]   打印带名字的排序\n"
    "  uv run rating.py --find <ID或名字>          定位某个道具\n"
    "  uv run rating.py --dump [文件名]            导出当前进度 xlsx\n"
    "  uv run rating.py --rebuild [xlsx]           从 xlsx 回灌排名（可先改再回灌）\n"
    "全部评完后会进入「校正模式」，可查看/移动/交换。"
)


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
    if not os.path.exists(STATE_FILE):
        return None
    with open(STATE_FILE, 'r', encoding='utf-8') as f:
        state = json.load(f)
    # 兼容早期版本的状态文件：补齐后加的字段
    state.setdefault('rerate_id', None)
    state.setdefault('rerate_low', 0)
    state.setdefault('rerate_high', 0)
    return state


def new_state():
    return {
        'sorted_ids': [],       # 已排好序的道具（从低到高）
        'next_index': 0,        # 下一个待处理的道具在 items 中的下标
        'current_item_id': None,  # 正在插入的道具
        'low': 0,
        'high': 0,
        'rerate_id': None,      # 正在「重新评级」的道具（功能 B）
        'rerate_low': 0,
        'rerate_high': 0,
    }


def load_all_names(excel_path):
    """读取所有道具的 ID -> 中文名（包含被 -1 排除的），只用于提示信息"""
    try:
        df = pd.read_excel(excel_path, sheet_name=SHEET_NAME)
    except Exception:
        return {}
    names = {}
    for _, row in df.iterrows():
        try:
            names[int(row['道具ID'])] = str(row['中文名'])
        except (TypeError, ValueError, KeyError):
            continue
    return names


def sanitize_state(state, item_dict, total, names=None):
    """清理状态中已经不在道具表里的 ID。

    源表是可能被改的（例如把某个道具的里该隐评级改成 -1 排除掉），
    这时状态文件里还留着它的 ID，直接用会 KeyError。
    """
    names = names or {}
    changed = False

    def label(item_id):
        return f"{item_id}" + (f"（{names[item_id]}）" if item_id in names else "")

    dropped = [i for i in state['sorted_ids'] if i not in item_dict]
    if dropped:
        state['sorted_ids'] = [i for i in state['sorted_ids'] if i in item_dict]
        changed = True
        shown = "、".join(label(i) for i in dropped[:10])
        more = "…" if len(dropped) > 10 else ""
        print(f"注意：{len(dropped)} 个道具已不在道具表里（可能被你改成 -1 了），"
              f"已从排序中移除：{shown}{more}")

    if state.get('rerate_id') is not None and state['rerate_id'] not in item_dict:
        print(f"注意：正在重新评级的道具 {label(state['rerate_id'])} 已不在道具表里，"
              f"已取消该会话。")
        state['rerate_id'] = None
        state['rerate_low'] = 0
        state['rerate_high'] = 0
        changed = True

    if state['current_item_id'] is not None and \
            state['current_item_id'] not in item_dict:
        print(f"注意：正在插入的道具 {label(state['current_item_id'])} 已不在道具表里，"
              f"已跳过它。")
        state['current_item_id'] = None
        state['low'] = 0
        state['high'] = 0
        changed = True

    # 列表变短后，比较区间可能越界，夹回合法范围
    n = len(state['sorted_ids'])
    if state['low'] > n:
        state['low'] = n
        changed = True
    if state['high'] >= n:
        state['high'] = n - 1
        changed = True
    if state['rerate_low'] > n:
        state['rerate_low'] = n
        changed = True
    if state['rerate_high'] >= n:
        state['rerate_high'] = n - 1
        changed = True
    if state['next_index'] > total:
        state['next_index'] = total
        changed = True

    if changed:
        save_state(state)
    return changed


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


# ---------------- 查看排序（把 ID 翻成中文名） ----------------

def print_ranking(state, item_dict, start=1, count=None, mark_ids=()):
    """紧凑打印排序：名次 / ID / 中文名 / 品质；mark_ids 里的道具标出来"""
    ids = state['sorted_ids']
    if not ids:
        print("（当前列表为空）")
        return
    if count is None:
        count = len(ids)
    total = len(ids)
    start = max(1, start)
    end = min(start - 1 + max(1, count), total)
    if start - 1 >= total:
        print(f"起始名次超出范围（当前共 {total} 个）。")
        return
    for i in range(start - 1, end):
        item = item_dict.get(ids[i])
        if item is None:
            print(f"{i + 1:>4}/{total}  {ids[i]:>4}  (不在道具表里)")
            continue
        mark = "  <== 它" if ids[i] in mark_ids else ""
        print(f"{i + 1:>4}/{total}  {item['id']:>4}  {item['中文名']}"
              f"（品质 {item['品质']}）{mark}")


# ---------------- 道具查找 ----------------

def resolve_item(query, items, item_dict):
    """把「道具 ID」或「名字片段」解析成一个道具；返回 (item, 候选列表)"""
    q = (query or '').strip()
    if not q:
        return None, []
    if q.isdigit():
        item = item_dict.get(int(q))
        return (item, []) if item else (None, [])
    ql = q.lower()
    matches = [it for it in items
               if q in it['中文名'] or ql in it['英文名'].lower()]
    if len(matches) == 1:
        return matches[0], []
    return None, matches


def lookup(query, items, item_dict):
    """解析道具，失败时打印候选/提示，返回 item 或 None"""
    item, cands = resolve_item(query, items, item_dict)
    if item:
        return item
    if cands:
        print(f"匹配到 {len(cands)} 个，请改用 ID 指定：")
        for it in cands[:15]:
            print(f"   {it['id']:>4}  {it['中文名']}")
    else:
        print(f"没有找到匹配「{query}」的道具。")
    return None


def ask_item(items, item_dict, tip):
    """循环询问，直到拿到一个道具；回车或 q 取消"""
    while True:
        raw = input(tip).strip()
        if raw == '' or raw.lower() == 'q':
            return None
        item = lookup(raw, items, item_dict)
        if item:
            return item


def ask_rank(max_rank):
    while True:
        raw = input(f"目标名次（1 ~ {max_rank}，回车取消）：").strip()
        if raw == '' or raw.lower() == 'q':
            return None
        if raw.isdigit():
            return int(raw)
        print("请输入数字名次。")


def report_pos(state, item_dict, item_id, pos, prefix):
    it = item_dict.get(item_id)
    name = it['中文名'] if it else '(不在道具表里)'
    print(f"\n>>> {prefix}：{name}（ID: {item_id}）"
          f"→ 第 {pos + 1} 名（当前共 {len(state['sorted_ids'])} 个）")


# ---------------- 功能 C：手动指定名次 ----------------

def manual_insert(state, items, item_dict, item_id, rank):
    """把 item_id 放到第 rank 名（1 起）。若它正是当前会话的道具，则结束该会话。"""
    sorted_ids = state['sorted_ids']
    moved = item_id in sorted_ids
    if moved:
        sorted_ids.remove(item_id)
    pos = max(0, min(rank - 1, len(sorted_ids)))
    sorted_ids.insert(pos, item_id)

    if state['current_item_id'] == item_id:
        state['next_index'] += 1
        state['current_item_id'] = None
        state['low'] = 0
        state['high'] = 0
    if state['rerate_id'] == item_id:
        state['rerate_id'] = None
        state['rerate_low'] = 0
        state['rerate_high'] = 0
    save_state(state)
    report_pos(state, item_dict, item_id, pos, "已移动" if moved else "已插入")


# ---------------- 功能 B：重新评级 ----------------

def start_rerate(state, item_dict, item_id):
    """把已排过名的道具摘出来，重新走一遍比较"""
    sorted_ids = state['sorted_ids']
    if state['rerate_id'] is not None:
        other = item_dict[state['rerate_id']]['中文名']
        print(f"正在进行「{other}」的重新评级，先完成它（或用 i 直接指定名次）。")
        return
    if state['current_item_id'] == item_id:
        # 正在插入的就是它：把比较区间重置为全部
        state['low'] = 0
        state['high'] = len(sorted_ids) - 1
        save_state(state)
        print(f"\n「{item_dict[item_id]['中文名']}」正在插入中，"
              f"已重置比较区间，重新开始比较。")
        return
    if item_id not in sorted_ids:
        print("这个道具还没排过名（或还没轮到），无法重新评级。")
        return
    sorted_ids.remove(item_id)
    state['rerate_id'] = item_id
    state['rerate_low'] = 0
    state['rerate_high'] = len(sorted_ids) - 1
    save_state(state)
    print(f"\n开始重新评级「{item_dict[item_id]['中文名']}」，"
          f"已有 {len(sorted_ids)} 个道具参与比较。")


# ---------------- 二分比较 ----------------

def compare_step(state, items, item_dict, which):
    """推进一次比较；返回 'continue'（或 'quit' 表示用户要退出）"""
    sorted_ids = state['sorted_ids']
    if which == 'rerate':
        item_id = state['rerate_id']
        k_low, k_high = 'rerate_low', 'rerate_high'
        title = "待重新评级道具："
    else:
        item_id = state['current_item_id']
        k_low, k_high = 'low', 'high'
        title = "待插入道具："
    low, high = state[k_low], state[k_high]

    # 源表被改动过时，正在处理的这个道具可能已经不在道具表里了
    if item_id not in item_dict:
        print(f"注意：道具 {item_id} 已不在道具表里，跳过这个会话。")
        if which == 'rerate':
            state['rerate_id'] = None
            state['rerate_low'] = 0
            state['rerate_high'] = 0
        else:
            state['current_item_id'] = None
            state['low'] = 0
            state['high'] = 0
        save_state(state)
        return 'continue'

    # 列表变短后比较区间可能越界，夹回合法范围
    if high >= len(sorted_ids):
        high = len(sorted_ids) - 1
        state[k_high] = high

    # 找到位置 → 插入
    if low > high:
        pos = low
        sorted_ids.insert(pos, item_id)
        if which == 'rerate':
            state['rerate_id'] = None
            state['rerate_low'] = 0
            state['rerate_high'] = 0
            save_state(state)
            report_pos(state, item_dict, item_id, pos, "已重新插入")
        else:
            state['next_index'] += 1
            state['current_item_id'] = None
            state['low'] = 0
            state['high'] = 0
            save_state(state)
            report_pos(state, item_dict, item_id, pos, "已插入")
            print(f"当前进度：{len(sorted_ids)} / {len(items)}")
        return 'continue'

    mid = (low + high) // 2
    mid_id = sorted_ids[mid]
    current_item = item_dict.get(item_id)
    mid_item = item_dict.get(mid_id)
    if mid_item is None:
        # 排序列表里混进了已不在道具表里的 ID（源表被改过），剔除后重来
        print(f"注意：排序列表里的 {mid_id} 已不在道具表里，已移除。")
        sorted_ids.pop(mid)
        save_state(state)
        return 'continue'
    if current_item is None:
        print(f"注意：道具 {item_id} 已不在道具表里，跳过这个会话。")
        state['current_item_id'] = None
        state['low'] = 0
        state['high'] = 0
        save_state(state)
        return 'continue'

    print("\n" + "=" * 60)
    print(title)
    print(format_item(current_item))
    print("\n" + "-" * 60)
    print("已排序道具（中间位置）：")
    print(format_item(mid_item))
    print("=" * 60)
    print(f"比较区间：第 {low + 1} ~ {high + 1} 名")
    print("哪个道具的里该隐评级更好（更高）？")
    print("输入 1：待插入道具更好")
    print("输入 2：已排序道具更好")
    print("输入 =：两者相等（按 ID 升序）")
    print("输入 l：查看已排序清单（写 l 起 数 可只看一段）")
    print("输入 f：查某个道具排在第几名")
    print("输入 r：重新评级某个道具")
    print("输入 i：手动指定名次（自由插入）")
    print("输入 q：保存并退出")
    choice = input("你的选择：").strip().lower()

    if choice == 'q':
        save_state(state)
        return 'quit'

    if choice == 'l' or choice.startswith('l '):
        nums = [p for p in choice.split()[1:] if p.isdigit()]
        start = int(nums[0]) if nums else 1
        count = int(nums[1]) if len(nums) > 1 else None
        print()
        print_ranking(state, item_dict, start, count)
        if count is None:
            print("（只想看一段可以写：l 100 20）")
        return 'continue'
    elif choice.startswith('f '):
        it = lookup(choice[2:].strip(), items, item_dict)
        if it is not None:
            if it['id'] in sorted_ids:
                idx = sorted_ids.index(it['id'])
                print()
                print_ranking(state, item_dict, max(1, idx - 2), 5,
                              mark_ids={it['id']})
            else:
                print(f"「{it['中文名']}」不在已排序列表里（可能还没轮到）。")
        return 'continue'
    elif choice == '1':
        state[k_low] = mid + 1
    elif choice == '2':
        state[k_high] = mid - 1
    elif choice == '=':
        if item_id < mid_id:
            state[k_high] = mid - 1
        else:
            state[k_low] = mid + 1
    elif choice == 'r':
        target = ask_item(items, item_dict, "要重新评级的道具（ID 或名字，回车取消）：")
        if target:
            start_rerate(state, item_dict, target['id'])
        return 'continue'
    elif choice == 'i':
        raw = input(f"要指定名次的道具（ID 或名字，回车 = 当前这个"
                    f"「{current_item['中文名']}」）：").strip()
        if raw == '':
            target = current_item
        else:
            target = lookup(raw, items, item_dict)
            if target is None:
                return 'continue'
        rank = ask_rank(len(sorted_ids) + 1)
        if rank is not None:
            manual_insert(state, items, item_dict, target['id'], rank)
        return 'continue'
    else:
        print("无效输入，请重新输入。")
        return 'continue'

    save_state(state)
    return 'continue'


# ---------------- 功能 D：校正模式 ----------------

def correction_mode(state, items, item_dict):
    sorted_ids = state['sorted_ids']
    print("\n" + "=" * 60)
    print("进入校正模式")
    print("  show [起始名次] [数量]   查看排序（省略参数 = 全部）")
    print("  find <ID或名字>          查看某个道具的名次与前后邻近项")
    print("  move <ID或名字> <名次>   把道具移动到指定名次")
    print("  swap <A> <B>             交换两个道具的名次（建议用 ID）")
    print("  export                   立即写出 xlsx")
    print("  done / 回车              结束校正并写出 xlsx")
    print("=" * 60)

    while True:
        raw = input("校正> ").strip()
        if raw == '' or raw.lower() in ('done', 'd', 'q', 'quit', 'exit'):
            return

        parts = raw.split()
        cmd = parts[0].lower()

        if cmd in ('help', 'h', '?', 'help'):
            print("  show [起始名次] [数量] / find <道具> / move <道具> <名次> / "
                  "swap <A> <B> / export / done")

        elif cmd == 'show':
            start = int(parts[1]) if len(parts) > 1 else 1
            count = int(parts[2]) if len(parts) > 2 else None
            print_ranking(state, item_dict, start, count)

        elif cmd == 'find':
            query = " ".join(parts[1:])
            it = lookup(query, items, item_dict)
            if it is None:
                continue
            if it['id'] not in sorted_ids:
                print(f"「{it['中文名']}」当前不在排序列表里。")
                continue
            idx = sorted_ids.index(it['id'])
            print_ranking(state, item_dict, max(1, idx - 2), 5,
                          mark_ids={it['id']})

        elif cmd in ('move', 'insert', 'i'):
            if len(parts) < 3:
                print("用法：move <ID或名字> <名次>")
                continue
            query = " ".join(parts[1:-1])
            if not parts[-1].isdigit():
                print("名次要写成数字，例如：move 182 12")
                continue
            it = lookup(query, items, item_dict)
            if it is None:
                continue
            manual_insert(state, items, item_dict, it['id'], int(parts[-1]))

        elif cmd == 'swap':
            if len(parts) != 3:
                print("用法：swap <A> <B>（建议直接用道具 ID）")
                continue
            a = lookup(parts[1], items, item_dict)
            b = lookup(parts[2], items, item_dict)
            if a is None or b is None:
                continue
            if a['id'] not in sorted_ids or b['id'] not in sorted_ids:
                print("两个道具都必须在当前排序列表里。")
                continue
            ia = sorted_ids.index(a['id'])
            ib = sorted_ids.index(b['id'])
            sorted_ids[ia], sorted_ids[ib] = sorted_ids[ib], sorted_ids[ia]
            save_state(state)
            print(f"已交换：第 {ia + 1} 名「{a['中文名']}」 ↔ "
                  f"第 {ib + 1} 名「{b['中文名']}」")

        elif cmd == 'export':
            export_xlsx(state, item_dict)

        else:
            print("未知命令，输入 help 查看用法。")


def export_xlsx(state, item_dict, path=None):
    path = path or OUTPUT_FILE
    result_data = []
    for rank, item_id in enumerate(state['sorted_ids'], start=1):
        item = item_dict.get(item_id)
        if item is None:
            print(f"注意：ID {item_id} 不在道具表里，导出时已跳过。")
            continue
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
    result_df.to_excel(path, index=False)
    print(f"已保存至：{path}（共 {len(result_df)} 个道具）")


def load_ranks_from_xlsx(path, item_dict):
    """从之前导出的 xlsx 读回排名（可在 Excel 里直接改顺序）"""
    if not os.path.exists(path):
        print(f"找不到文件：{path}")
        return None
    df = pd.read_excel(path)
    rank_col, id_col = '排名（从低到高）', '道具ID'
    if rank_col not in df.columns or id_col not in df.columns:
        print(f"xlsx 需要包含「{rank_col}」和「{id_col}」两列。")
        return None
    df = df.sort_values(rank_col)
    ids, seen, skipped = [], set(), 0
    for v in df[id_col]:
        try:
            item_id = int(v)
        except (TypeError, ValueError):
            continue
        if item_id in seen:
            continue
        seen.add(item_id)
        if item_id not in item_dict:
            skipped += 1
            continue
        ids.append(item_id)
    if skipped:
        print(f"注意：{skipped} 个 ID 不在道具表里，已忽略。")
    return ids


# ---------------- 主流程 ----------------

def main():
    args = sys.argv[1:]

    def arg_value(flag, default=''):
        """取 --flag 后面那个参数；没写值就返回 default；flag 不存在返回 None"""
        if flag not in args:
            return None
        i = args.index(flag)
        if len(args) > i + 1 and not args[i + 1].startswith('--'):
            return args[i + 1]
        return default

    rebuild_path = arg_value('--rebuild', OUTPUT_FILE)
    dump_path = arg_value('--dump', PROGRESS_FILE)
    find_query = arg_value('--find', '')
    show_start = show_count = None
    if '--show' in args:
        nums = [a for a in args[args.index('--show') + 1:] if a.isdigit()]
        show_start = int(nums[0]) if nums else 1
        show_count = int(nums[1]) if len(nums) > 1 else None

    print("=== 里该隐道具评级（两两比较，从低到高，选择更好的） ===")

    if not os.path.exists(EXCEL_FILE):
        print(f"找不到文件：{EXCEL_FILE}")
        excel_path = input("请输入 Excel 文件路径：").strip()
        if not os.path.exists(excel_path):
            print("文件不存在，退出。")
            return
    else:
        excel_path = EXCEL_FILE

    print("\n正在加载道具数据...")
    items = load_items(excel_path)
    item_dict = {item['id']: item for item in items}
    total = len(items)
    print(f"共加载 {total} 个可合成道具（已排除里该隐评级为 -1 的道具）。")

    # ---------- 只读模式：--show / --find / --dump（不进入比较） ----------
    if show_start is not None or find_query is not None or dump_path is not None:
        state = load_state()
        if state is None:
            print("没有找到状态文件，先跑一次评级（不加参数）再来查看。")
            return
        sanitize_state(state, item_dict, total, load_all_names(excel_path))
        if show_start is not None:
            print()
            print_ranking(state, item_dict, show_start, show_count)
        if find_query is not None:
            it = lookup(find_query, items, item_dict)
            if it is not None:
                if it['id'] in state['sorted_ids']:
                    idx = state['sorted_ids'].index(it['id'])
                    print_ranking(state, item_dict, max(1, idx - 2), 5,
                                  mark_ids={it['id']})
                else:
                    print(f"「{it['中文名']}」不在已排序列表里（可能还没轮到）。")
        if dump_path is not None:
            export_xlsx(state, item_dict, dump_path)
        print(f"\n已排序 {len(state['sorted_ids'])} / {total}；"
              f"道具表下一个待处理的是第 {state['next_index'] + 1} 个")
        if state.get('current_item_id'):
            it = item_dict.get(state['current_item_id'])
            name = it['中文名'] if it else state['current_item_id']
            print(f"当前正在插入：{name}"
                  f"（比较区间第 {state['low'] + 1} ~ {state['high'] + 1} 名）")
        print("改完顺序后可用 --rebuild 回灌；直接重跑不带参数可继续评级。")
        return

    print(HELP_TIP)

    # ---------- 从已有 xlsx 回灌排名 ----------
    if rebuild_path is not None:
        sorted_ids = load_ranks_from_xlsx(rebuild_path, item_dict)
        if sorted_ids is None:
            return
        state = new_state()
        state['sorted_ids'] = sorted_ids
        state['next_index'] = 0   # 主循环会自动跳过已排名的道具
        save_state(state)
        print(f"已从 {rebuild_path} 读回 {len(sorted_ids)} 个道具的排名。")
        if len(sorted_ids) < total:
            print(f"还有 {total - len(sorted_ids)} 个道具没排名，接着比较。")
    else:
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
            state = new_state()
            save_state(state)
        else:
            # 源表可能被改过（例如把道具改成 -1 排除掉），先清理状态
            sanitize_state(state, item_dict, total, load_all_names(excel_path))

    while True:
        sorted_ids = state['sorted_ids']

        # 1) 重新评级会话优先处理
        if state.get('rerate_id') is not None:
            if compare_step(state, items, item_dict, 'rerate') == 'quit':
                print("已保存，退出。")
                return
            continue

        # 2) 取下一个待插入的道具：按源表顺序找第一个还没排名的
        #    （不依赖历史下标，这样在源表里增删/调整行、或把 -1 改回正常值都不会错位）
        if state['current_item_id'] is None:
            ranked = set(sorted_ids)
            next_idx = None
            for k, it in enumerate(items):
                if it['id'] not in ranked:
                    next_idx = k
                    break
            if next_idx is None:
                print("\n所有道具已评级完成！")
                break
            state['next_index'] = next_idx
            current_item = items[next_idx]
            state['current_item_id'] = current_item['id']
            state['low'] = 0
            state['high'] = len(sorted_ids) - 1
            save_state(state)
            continue

        # 3) 正常的二分比较
        if compare_step(state, items, item_dict, 'main') == 'quit':
            print("已保存，退出。")
            return

    # —— 收尾：校正模式 + 导出 ——
    correction_mode(state, items, item_dict)
    export_xlsx(state, item_dict)
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        print("状态文件已删除。")


if __name__ == '__main__':
    main()
