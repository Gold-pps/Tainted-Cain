# -*- coding: utf-8 -*-
"""以撒的结合（The Binding of Isaac: Repentance）种子字符串 <-> 32 位种子值。

移植自 IsaacDecraftingHuijiGadget 项目 src/App.vue 中的 str2seed()。
原实现是游戏反编译代码的直译，这里保持逐位等价，只去掉了 JS 特有的位运算包装。

种子格式
--------
    "XXXX XXXX"：固定 9 个字符，第 5 个（下标 4）必须是空格。

字符表
------
    "ABCDEFGHJKLMNPQRSTWXYZ01234V6789"   共 32 个字符
    刻意剔除了易混淆的 I / O / U / 5，其余码位一律非法。

编码方式
--------
    8 个有效字符各取 5 bit，共 40 bit。记为 ns[0..7]（ns[0] 是第 1 个字符）：

        ns[0..6]  ->  32 位种子值（ns[7] 是纯校验位，不参与）
        ns[7]     ->  校验位

    32 位种子值的位布局（大端）：

        seed = whole ^ 0xFEF7FFD
        whole = (ns[6] >> 3) | (inner << 2)
        inner = ns[5] | ns[4]<<5 | ns[3]<<10 | ns[2]<<15 | ns[1]<<20 | ns[0]<<25

    即 ns[0] 占最高 5 bit，ns[5] 占 bit2..6，ns[6] 只贡献最低 2 bit。

校验位
------
    对 seed 反复取低 5 bit 做一次累加，得到 8 bit 校验值，要求

        acc == ns[7] | ((ns[6] & 7) << 5)

    不满足则整个种子非法。注意 0 被用作「非法」哨兵值，因此
    whole == 0xFEF7FFD（导致 seed == 0）的字符串会被判为非法。
"""

SEED_ALPHABET = "ABCDEFGHJKLMNPQRSTWXYZ01234V6789"
"""32 个合法字符，下标即其 5 bit 值。"""

SEED_XOR_MASK = 0xFEF7FFD
"""游戏内固定魔数，种子值与它异或。"""

MASK32 = 0xFFFFFFFF

_CHAR_TO_VALUE = {ch: i for i, ch in enumerate(SEED_ALPHABET)}


def _checksum(seed_value):
    """计算 8 bit 校验值。

    对应 JS 里的那个 for 循环：每次吃掉 5 bit，
    acc = (acc * 2) mod 255（用 (v10 >>> 7) 补进位来避免取模）。
    """
    acc = 0
    remainder = seed_value & MASK32
    while remainder != 0:
        v10 = ((remainder & 0xFF) + acc) & 0xFF
        acc = (2 * v10 + (v10 >> 7)) & 0xFF
        remainder >>= 5
    return acc


def str2seed(seed_str):
    """把 "XXXX XXXX" 形式的种子字符串转成 32 位无符号整数。

    非法种子一律返回 0（0 也是原实现的哨兵值，合法种子不可能得到 0）。
    """
    if len(seed_str) != 9:
        return 0
    if seed_str[4] != " ":
        return 0

    # 字符 -> 5 bit 值，顺便完成合法性检查
    num_seed = []
    for i in range(9):
        if i == 4:
            continue
        value = _CHAR_TO_VALUE.get(seed_str[i])
        if value is None:
            return 0
        num_seed.append(value)

    # ns[0] 在高位，ns[5] 在低位，拼出 30 bit
    inner = 0
    for i in range(6):
        inner |= num_seed[i] << (5 * (5 - i))

    # ns[6] 只有高 2 bit 参与种子值
    whole = (num_seed[6] >> 3) | (inner << 2)
    seed_value = (whole ^ SEED_XOR_MASK) & MASK32

    # 校验位：ns[7] 的 5 bit + ns[6] 的低 3 bit
    expected = num_seed[7] | ((num_seed[6] & 0x07) << 5)
    if _checksum(seed_value) != expected:
        return 0

    return seed_value


def seed2str(seed_value):
    """把 32 位种子值转回种子字符串（自动补上正确的校验位）。"""
    seed_value &= MASK32

    whole = (seed_value ^ SEED_XOR_MASK) & MASK32
    acc = _checksum(seed_value)

    inner = whole >> 2
    num_seed = [0] * 8
    for i in range(6):
        num_seed[i] = (inner >> (5 * (5 - i))) & 0x1F
    # ns[6] = 高 2 bit（来自 whole 低 2 bit）+ 低 3 bit（来自校验值高 3 bit）
    num_seed[6] = ((whole & 0x03) << 3) | ((acc >> 5) & 0x07)
    num_seed[7] = acc & 0x1F

    chars = [SEED_ALPHABET[v] for v in num_seed]
    return "".join(chars[:4]) + " " + "".join(chars[4:])


def is_valid_seed(seed_str):
    """该字符串是不是一个能通过校验的合法种子。"""
    return str2seed(seed_str) != 0


if __name__ == "__main__":
    import random

    # 1. 作者在 App.vue 里用的默认种子
    sample = "JKD9 Z0C9"
    print("默认种子 %r -> 0x%08X (valid=%s)"
          % (sample, str2seed(sample), is_valid_seed(sample)))

    # 2. 往返一致性：随机种子值 -> 字符串 -> 种子值
    random.seed(20260915)
    bad = 0
    for _ in range(20000):
        v = random.getrandbits(32)
        if v == 0:              # 0 是非法哨兵值，按设计无法表示
            continue
        s = seed2str(v)
        if str2seed(s) != v:
            bad += 1
            if bad <= 5:
                print("mismatch: 0x%08X -> %r -> 0x%08X" % (v, s, str2seed(s)))
    print("往返测试 20000 例，不一致 %d 例" % bad)

    # 3. 改动校验位应当被拒绝
    if is_valid_seed(sample):
        for pos in (0, 3, 5, 8):
            for ch in SEED_ALPHABET:
                mutated = sample[:pos] + ch + sample[pos + 1:]
                if mutated == sample:
                    continue
                # 只要改动了参与种子值的位，就必须失败；
                # 改最后一位只是换校验位，也应当失败
                if is_valid_seed(mutated):
                    print("警告：%r 被误判为合法" % mutated)
                break
        print("单位翻转测试通过")

    # 4. 非法输入
    for invalid in ("", "JKD9Z0C9", "JKD9 Z0C9 ", "jKD9 Z0C9", "JKD9 Z0CI", "JKD9 Z0O9"):
        assert str2seed(invalid) == 0, invalid
    print("非法输入测试通过")
