-- Crafting Recorder —— 自动记录里该隐合成宝袋配方与合成结果（需要 REPENTOGON）
--
-- 工作原理：
--   合成袋装满 8 个掉落物时立即记录「配方 → 预期产物」（参考 EID 的"已知配方"，
--   不必真的挥袋合成）；袋子被清空时再兜底记录一次，同一配方会被去重。
--   记录追加写入 ModData（Isaac.SaveModData），供外部 Python 工具读取。
--   为了不影响游戏体验，记录时不显示任何游戏内提示。
--
-- 记录格式：配方|道具ID|道具中文名|种子|掉落物价值之和|是否合成|十字圣球|本局合成次数
--   道具中文名来自同目录的 item_names_zh.lua（由《以撒的结合忏悔+_全道具信息表.xlsx》
--   生成）。游戏内 ItemConfig.Name 只会给出 #XXX_NAME 这类语言键，所以要靠该表翻译；
--   表里没有的 ID 才会退回游戏内的英文名/语言键。
--   掉落物价值之和 = Σ 该掉落物数量 × 单件价值（价值表见 PICKUP_VALUES）。
--   是否合成：凑满袋子即记录为“否”，真的挥袋合成后原地更新为“是”。
--   十字圣球：记录该配方时角色是否持有十字圣球（会改变合成袋产出的品质）。
--   若持有圣球后同一配方产出变了，会按新的产物 ID 单独记一行；产出没变则原地更新该标记。
--   本局合成次数：该配方在这一局（同一存档、同一种子）被真正搓出来的次数。
--   只是凑满袋子不算，每挥袋合成一次 +1；按「种子+配方+产物」各自计数 ——
--   圣球把产物改变时，两种产物各有自己的次数。
--
-- 存档位置（Windows，注意在游戏安装目录，不是"我的文档"）：
--   <游戏目录>\data\crafting_recorder\save<存档槽位>.dat
--   例如 C:\Program Files (x86)\Steam\steamapps\common\
--        The Binding of Isaac Rebirth\data\crafting_recorder\save2.dat
--
-- 重要：Isaac.HasModData / LoadModData 只有在选定存档槽位后才有效。
--   因此历史记录的载入放在 MC_POST_GAME_STARTED 以及首次记录之前，
--   绝不能在 main.lua 加载阶段直接调用；否则读不到旧记录，
--   新记录落盘时会整份覆盖，导致上次游戏的记录丢失。

if not REPENTOGON then
    Isaac.DebugString("[CraftingRecorder] 未检测到 REPENTOGON，模组停用")
    return
end

local mod = RegisterMod("Crafting Recorder", 1)

-- BagOfCraftingPickup 值 1..28 对应的中文名
-- （与《合成宝袋组件.xlsx》的“顺序”列一致；29 = 便便，不参与记录）
local PICKUP_NAMES = {
    [1] = "红心", [2] = "魂心", [3] = "黑心", [4] = "永恒之心",
    [5] = "金心", [6] = "骨心", [7] = "腐心",
    [8] = "硬币", [9] = "镍币", [10] = "铸币", [11] = "幸运硬币",
    [12] = "钥匙", [13] = "金钥匙", [14] = "充能钥匙",
    [15] = "炸弹", [16] = "金炸弹", [17] = "巨型炸弹",
    [18] = "微型电池", [19] = "小电池", [20] = "超级电池",
    [21] = "卡牌", [22] = "胶囊", [23] = "符文/魂石",
    [24] = "骰子碎片", [25] = "红钥匙碎片",
    [26] = "金硬币", [27] = "金胶囊", [28] = "金电池",
}

-- 每件掉落物的价值/品质（来源：《合成宝袋组件.xlsx》的“品质”列，键为“顺序”列）
local PICKUP_VALUES = {
    [1] = 1, [2] = 4, [3] = 5, [4] = 5, [5] = 5, [6] = 5, [7] = 1,
    [8] = 1, [9] = 3, [10] = 5, [11] = 8,
    [12] = 2, [13] = 7, [14] = 5,
    [15] = 2, [16] = 7, [17] = 10,
    [18] = 2, [19] = 4, [20] = 8,
    [21] = 2, [22] = 2, [23] = 4, [24] = 4, [25] = 2,
    [26] = 7, [27] = 7, [28] = 7,
}

-- 道具 ID -> 中文名（自动生成，与《合成宝袋组件.xlsx》同源的参考表一致）
local ZH_NAMES
do
    local ok, t = pcall(require, "item_names_zh")
    if ok and type(t) == "table" then
        ZH_NAMES = t
    else
        Isaac.DebugString(
            "[CraftingRecorder] 未载入 item_names_zh.lua，道具名将退回英文")
    end
end

local state = {
    prev = {},          -- 上一帧的袋内容（1..8）
    wasFull = false,    -- 上一帧袋子是否已满
    pendingOutput = 0,  -- 袋子满时记录的预期产物
    wasOrb = false,     -- 上一帧是否持有十字圣球
    lines = {},         -- 已捕获的记录行（持久化到 ModData）
    lineIndex = {},     -- 记录行 -> 在 lines 中的下标（用于原地更新）
    keyLine = {},       -- seed|recipe|output -> 当前对应的记录行
    craftCount = {},    -- seed|recipe|output -> 本局已合成次数
    loaded = false,     -- 是否已成功读到过 ModData
    loadTries = 0,      -- 载入尝试次数（时机未到时允许重试，但有上限）
}

local MAX_LOAD_TRIES = 5

-- ---------- 小工具 ----------

local function snapshot(content)
    local t = {}
    for i = 1, 8 do
        t[i] = content[i] or 0
    end
    return t
end

local function sameBag(a, b)
    for i = 1, 8 do
        if (a[i] or 0) ~= (b[i] or 0) then
            return false
        end
    end
    return true
end

-- 把袋内容标准化成「红心2+硬币3+钥匙1」格式（按枚举值升序，与 Excel 顺序一致）
local function buildRecipe(bag)
    local counts = {}
    for i = 1, 8 do
        local v = bag[i] or 0
        if v >= 1 and v <= 28 then
            counts[v] = (counts[v] or 0) + 1
        end
    end
    local parts = {}
    for v = 1, 28 do
        if counts[v] then
            parts[#parts + 1] = PICKUP_NAMES[v] .. counts[v]
        end
    end
    return table.concat(parts, "+")
end

-- 掉落物价值之和
local function buildValue(bag)
    local total = 0
    for i = 1, 8 do
        local v = bag[i] or 0
        total = total + (PICKUP_VALUES[v] or 0)
    end
    return total
end

local function getSeedString()
    local ok, s = pcall(function()
        return Game():GetSeeds():GetStartSeedString()
    end)
    if ok and s then
        return s
    end
    return ""
end

-- ---------- ModData 读写 ----------

local function saveData()
    Isaac.SaveModData(mod, table.concat(state.lines, "\n"))
end

-- 按 | 切分记录行（兼容旧格式：配方|ID|名字|种子）
local function splitFields(line)
    local f = {}
    for part in string.gmatch(line, "([^|]*)") do
        f[#f + 1] = part
    end
    return f
end

-- seed|recipe|output 是记录的稳定标识（与价值/是否合成两栏无关）
local function keyOfFields(f)
    if not f[1] or not f[2] or not f[4] then
        return nil
    end
    return f[4] .. "|" .. f[1] .. "|" .. f[2]
end

-- 合成次数与记录同键：种子+配方+产物

-- 把一行记录加入内存（整行去重），返回是否为新增
local function registerLine(line)
    if not line or line == "" then
        return false
    end
    if state.lineIndex[line] then
        return false
    end
    state.lines[#state.lines + 1] = line
    state.lineIndex[line] = #state.lines

    local f = splitFields(line)
    local key = keyOfFields(f)
    if key then
        state.keyLine[key] = line
    end

    -- 从存档里恢复该记录的合成次数；旧记录没有这一栏时，
    -- 只要“是否合成”是“是”，至少可以确定合成过 1 次
    local n = tonumber(f[8])
    if not n and f[6] == "是" then
        n = 1
    end
    if key and n and n > (state.craftCount[key] or 0) then
        state.craftCount[key] = n
    end
    return true
end

-- 载入 ModData 中的历史记录（务必在选定存档槽位之后调用）
local function loadData()
    if state.loaded or state.loadTries >= MAX_LOAD_TRIES then
        return 0
    end
    state.loadTries = state.loadTries + 1

    local ok, raw = pcall(function()
        if not Isaac.HasModData(mod) then
            return ""
        end
        return Isaac.LoadModData(mod) or ""
    end)
    if not ok then
        return 0
    end

    if not raw or raw == "" then
        -- 可能确实没有历史记录，也可能是时机未到（未选存档槽位），留待下次重试
        return 0
    end
    state.loaded = true

    local n = 0
    for line in string.gmatch(raw, "[^\r\n]+") do
        if registerLine(line) then
            n = n + 1
        end
    end
    if n > 0 then
        Isaac.DebugString("[CraftingRecorder] 载入历史记录 " .. n .. " 条")
    end
    return n
end

local function ensureLoaded()
    if not state.loaded then
        loadData()
    end
end

-- ---------- 记录 ----------

-- 十字圣球：持有它会改变合成宝袋产出的品质，所以记录时要把状态标出来
local SACRED_ORB = CollectibleType.COLLECTIBLE_SACRED_ORB or 691

local function hasSacredOrb(player)
    if not player then
        return false
    end
    local ok, v = pcall(function()
        return player:HasCollectible(SACRED_ORB)
    end)
    return ok and v and true or false
end

-- 优先用中文名表，取不到才退回游戏内名字（通常是 #XXX_NAME 语言键）
local function itemNameOf(output)
    if ZH_NAMES and ZH_NAMES[output] then
        return ZH_NAMES[output]
    end
    local cfg = Isaac.GetItemConfig():GetCollectible(output)
    if cfg then
        return cfg.Name or ""
    end
    return ""
end

-- 写入/更新一条记录（按 seed|recipe|output 去重）
--   craftedEvent = false：刚凑满袋子，还没真的合成
--   craftedEvent = true ：确实挥袋合成了（同时把“本局合成次数” +1）
--   player 用于判断当前是否持有十字圣球
local function upsertRecord(bag, output, craftedEvent, player)
    if not output or output <= 0 then
        return false
    end

    local recipe = buildRecipe(bag)
    if recipe == "" then
        return false
    end

    -- 写入前先确保旧记录已载入，避免覆盖历史（重启游戏后丢记录的原因）
    ensureLoaded()

    local seedStr = getSeedString()
    local key = seedStr .. "|" .. recipe .. "|" .. output
    local old = state.keyLine[key]

    -- “是否合成”只会从否变成是；同一个配方再凑一次也不会回退成否
    local crafted = craftedEvent
    if old and splitFields(old)[6] == "是" then
        crafted = true
    end

    -- 本局合成次数：只有真的合成事件才 +1（凑满袋子不算），
    -- 按「种子+配方+产物」各自计数
    if craftedEvent then
        state.craftCount[key] = (state.craftCount[key] or 0) + 1
    end
    local count = state.craftCount[key] or 0

    local line = recipe .. "|" .. output .. "|" .. itemNameOf(output) .. "|"
        .. seedStr .. "|" .. buildValue(bag) .. "|" .. (crafted and "是" or "否")
        .. "|" .. (hasSacredOrb(player) and "是" or "否")
        .. "|" .. count

    if old == line then
        return false   -- 内容完全没有变化
    end

    if old then
        local idx = state.lineIndex[old]
        if idx then
            -- 原地更新（不新增行，例如合成完成后把“否”改成“是”）
            state.lines[idx] = line
            state.lineIndex[old] = nil
            state.lineIndex[line] = idx
            state.keyLine[key] = line
            saveData()
            Isaac.DebugString("[CraftingRecorder] 更新记录: " .. line)
            return true
        end
    end

    registerLine(line)
    saveData()

    -- 仅写日志文件，不在游戏内弹出任何提示
    Isaac.DebugString("[CraftingRecorder] 记录配方: " .. line)
    return true
end

-- ---------- 回调 ----------

mod:AddCallback(ModCallbacks.MC_POST_GAME_STARTED, function(_, isContinued)
    state.prev = {}
    state.wasFull = false
    state.pendingOutput = 0
    state.wasOrb = false
    state.loadTries = 0
    -- 此时存档槽位已确定，才能正确读取上一次游戏的记录
    loadData()
end)

-- 游戏内兜底：万一 POST_GAME_STARTED 时机太早没读到，进入游戏后再补几次
local loadCheckTimer = 0

mod:AddCallback(ModCallbacks.MC_POST_PEFFECT_UPDATE, function(_, player)
    -- 兜底载入旧记录（只在游戏内重试，此时存档槽位一定已确定）
    if not state.loaded then
        loadCheckTimer = loadCheckTimer + 1
        if loadCheckTimer >= 30 then
            loadCheckTimer = 0
            loadData()
        end
    end

    if not player:HasCollectible(CollectibleType.COLLECTIBLE_BAG_OF_CRAFTING) then
        state.prev = {}
        state.wasFull = false
        state.pendingOutput = 0
        state.wasOrb = false
        return
    end

    local snap = snapshot(player:GetBagOfCraftingContent())
    local count = 0
    for i = 1, 8 do
        if snap[i] ~= 0 then
            count = count + 1
        end
    end

    if count == 8 then
        -- 袋子已满：立即记录（此时“是否合成”= 否）
        -- 袋内容、预期产物、是否持有圣球，任意一项变化都要重新评估：
        -- 袋子满着的时候吃到圣球会改变产物，此时袋内容并没有变化。
        local output = player:GetBagOfCraftingOutput() or 0
        local orb = hasSacredOrb(player)
        local changed = not state.wasFull or not sameBag(state.prev, snap)
            or output ~= state.pendingOutput or orb ~= state.wasOrb
        if changed then
            upsertRecord(snap, output, false, player)
        end
        state.prev = snap
        state.wasFull = true
        state.pendingOutput = output
        state.wasOrb = orb

    elseif count == 0 and state.wasFull then
        -- 袋子从满变空：说明真的合成过了，把“是否合成”更新为“是”
        upsertRecord(state.prev, state.pendingOutput, true, player)
        state.wasFull = false
        state.pendingOutput = 0
        state.wasOrb = false

    elseif count > 0 and count < 8 then
        -- 装袋过程中，重置满袋跟踪
        state.wasFull = false
        state.pendingOutput = 0
        state.wasOrb = false
    end
end)

Isaac.DebugString("[CraftingRecorder] 模组已加载")
