const { Schema } = require('koishi')

// 词条表格列
const EntryValue = Schema.array(Schema.string()).default([]).description('回复内容列表（多个随机一条）')

const Config = Schema.object({
  library: Schema.string()
    .default('默认')
    .description('要管理的 word-core 词库名'),
  authorId: Schema.string()
    .default('')
    .description('操作者QQ号（用于 word-core API 调用，留空=用管理员QQ）'),
  entries: Schema.dict(EntryValue)
    .default({})
    .description('📊 词条表格（键=触发词，值=回复列表）。修改后保存即同步到 word-core'),
})

const inject = { required: ['word'] }

function apply(ctx, config) {
  // ===== 同步引擎 =====
  async function loadFromWordCore() {
    try {
      const db = config.library || '默认'
      const result = await ctx.word.editor.readWord(db)
      return result?.data || {}
    } catch (e) {
      ctx.logger.warn(`[word-manage] 读取词库 "${db}" 失败: ${e.message}`)
      return null
    }
  }

  async function syncConfigToWordCore(entries) {
    const db = config.library || '默认'
    const uid = config.authorId || 'admin'
    const current = await loadFromWordCore()
    if (!current) return false

    // 对比并同步
    const configKeys = new Set(Object.keys(entries))
    const coreKeys = new Set(Object.keys(current))
    let changed = false

    // 添加 / 更新
    for (const q of configKeys) {
      const newAnswers = entries[q] || []
      const oldAnswers = current[q] || []

      // 找出需要添加的答案
      for (const a of newAnswers) {
        if (!oldAnswers.includes(a)) {
          try {
            await ctx.word.editor.addWordItem(db, uid, q, a)
            ctx.logger.info(`[word-manage] + "${q}" → "${a}"`)
            changed = true
          } catch (e) {
            ctx.logger.error(`[word-manage] 添加失败 "${q}": ${e.message}`)
          }
        }
      }

      // 找出需要删除的答案
      if (oldAnswers.length > 0) {
        for (let i = oldAnswers.length - 1; i >= 0; i--) {
          if (!newAnswers.includes(oldAnswers[i])) {
            try {
              await ctx.word.editor.rmWordItem(db, uid, q, i)
              ctx.logger.info(`[word-manage] - "${q}" #${i}`)
              changed = true
            } catch (e) {
              ctx.logger.error(`[word-manage] 删除失败 "${q}" #${i}: ${e.message}`)
            }
          }
        }
      }
    }

    // 删除整个问题（移除所有答案）
    for (const q of coreKeys) {
      if (!configKeys.has(q)) {
        try {
          await ctx.word.editor.rmWordItem(db, uid, q, 'all')
          ctx.logger.info(`[word-manage] × 删除 "${q}"`)
          changed = true
        } catch (e) {
          ctx.logger.error(`[word-manage] 删除问题 "${q}" 失败: ${e.message}`)
        }
      }
    }

    return changed
  }

  // ===== 启动时加载 =====
  let entries = { ...(config.entries || {}) }
  let loaded = Object.keys(entries).length > 0

  ;(async () => {
    const fromCore = await loadFromWordCore()
    if (fromCore && !loaded) {
      // 设置页还没配过，从 word-core 加载
      entries = fromCore
      Object.assign(config.entries, entries)
      loaded = true
      ctx.logger.info(`[word-manage] 已从 word-core 加载 ${Object.keys(entries).length} 条词条`)
    } else if (loaded) {
      // 设置页有数据，同步到 word-core
      await syncConfigToWordCore(entries)
    }
  })()

  // ===== 命令 =====
  ctx.command('wmgr', '词库表格管理')
    .action(() =>
      '📊 表格管理：控制台 → 插件设置 → ll-word-manage → entries\n' +
      '💬 快速命令：\n' +
      'wmgr.list [关键词] — 查看词条\n' +
      'wmgr.add <触发词> <回复> — 添加词条\n' +
      'wmgr.rm <触发词> [序号] — 删除回复\n' +
      'wmgr.reload — 从 word-core 重新加载\n' +
      'wmgr.sync — 强制同步设置页 → word-core'
    )

  ctx.command('wmgr.list [keyword]', '查看词条列表 / 指定词条详情')
    .action(async ({ session }, keyword) => {
      // 实时从 word-core 读
      const data = await loadFromWordCore()
      if (!data) return '无法读取 word-core 词库，请检查 library 配置'
      const keys = Object.keys(data)
      if (!keys.length) return '词库为空'

      if (keyword) {
        const answers = data[keyword]
        if (!answers || !answers.length) return `未找到触发词 "${keyword}"`
        return `📋 "${keyword}" 的回复：\n${answers.map((a, i) => `  ${i}. ${a}`).join('\n')}`
      }

      const total = keys.reduce((sum, k) => sum + (data[k]?.length || 0), 0)
      const preview = keys.slice(0, 20).map(k => {
        const count = data[k]?.length || 0
        const first = data[k]?.[0] || ''
        return `  ${k} → ${first.substring(0, 25)}${first.length > 25 ? '...' : ''}${count > 1 ? ` (+${count - 1}条)` : ''}`
      }).join('\n')
      return `词库 "${config.library || '默认'}" 共 ${keys.length} 个触发词 / ${total} 条回复\n${preview}${keys.length > 20 ? `\n  ... 还有 ${keys.length - 20} 个` : ''}`
    })

  ctx.command('wmgr.add <question:string> <answer:text>', '添加词条（写入 word-core）')
    .action(async ({ session }, question, answer) => {
      const db = config.library || '默认'
      const uid = String(session.userId)
      if (!question || !answer) return '请指定触发词和回复内容'
      try {
        await ctx.word.editor.addWordItem(db, uid, question, answer)
        return `✅ "${question}" → "${answer}" 已添加到词库 "${db}"`
      } catch (e) {
        return `❌ 添加失败: ${e.message}`
      }
    })

  ctx.command('wmgr.rm <question:string> [index:number]', '删除回复（不填序号=删除整个触发词）')
    .action(async ({ session }, question, index) => {
      const db = config.library || '默认'
      const uid = String(session.userId)
      if (!question) return '请指定要删除的触发词'
      try {
        const which = index !== undefined ? index : 'all'
        await ctx.word.editor.rmWordItem(db, uid, question, which)
        return `✅ 已从词库 "${db}" 删除 "${question}"${index !== undefined ? ` #${index}` : '（全部）'}`
      } catch (e) {
        return `❌ 删除失败: ${e.message}`
      }
    })

  ctx.command('wmgr.reload', '从 word-core 重新加载所有词条')
    .action(async () => {
      const data = await loadFromWordCore()
      if (!data) return '❌ 加载失败'
      // 更新内存和 config
      for (const k of Object.keys(config.entries)) delete config.entries[k]
      Object.assign(config.entries, data)
      Object.assign(entries, data)
      return `✅ 已从 word-core 加载 ${Object.keys(data).length} 条词条`
    })

  ctx.command('wmgr.sync', '将设置页 entries 表格同步到 word-core')
    .action(async () => {
      const changed = await syncConfigToWordCore(config.entries || {})
      if (changed === false) return '❌ 同步失败，无法读取 word-core'
      return changed ? '✅ 已同步到 word-core' : '✅ 无需同步（数据一致）'
    })
}

module.exports = { Config, apply, inject }
