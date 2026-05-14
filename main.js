const { app, BrowserWindow, ipcMain, dialog, shell, Tray, Menu, nativeImage } = require('electron')
const path = require('path')
const { exec, execFile, spawn } = require('child_process')
const fs = require('fs')
const os = require('os')

// ─── Constants ───────────────────────────────────────────────────────────────
const IS_WIN = process.platform === 'win32'
const SETTINGS_PATH = path.join(app.getPath('userData'), 'jt_settings.json')
const LOG_PATH = path.join(app.getPath('userData'), 'jt_log.txt')

let mainWindow = null
let pulseWindow = null
let fivemSettingsWindow = null
let tray = null

// ─── Discord Rich Presence ────────────────────────────────────────────────────
const DISCORD_CLIENT_ID  = '1504472893818667152'
const DISCORD_INVITE     = 'https://discord.gg/jylligraffat'
const DISCORD_DOWNLOAD   = 'https://discord.com/channels/1295856827388399779/1295856828160282746'

let discordRPC            = null
let discordReady          = false
let discordStartTimestamp = null
let discordRetryTimer     = null

// State tracked by the renderer via IPC
let discordCurrentPage    = 'dashboard'
let discordFpsGained      = 0
let discordTweaksApplied  = 0
let discordHasFiveM       = false
let discordPulseGame      = null   // name of game Pulse is active for, or null
let discordActiveGame     = null   // game watcher detected game, or null

// Override state for active operations (null = idle, use page-based presence)
let discordActivityOverride = null   // { details, state, imageKey, imageText }
let discordOverrideCycleTimer = null // for cycling text during long operations

// Per-page presence config
const PAGE_PRESENCE = {
  'dashboard':        { details: 'On the Dashboard',          image: 'jylli_logo'  },
  'windows-opti':     { details: 'Tweaking Windows',          image: 'jylli_logo'  },
  'hardware':         { details: 'Tweaking Hardware',         image: 'jylli_logo'  },
  'network':          { details: 'Tweaking Network',          image: 'jylli_logo'  },
  'fivem':            { details: 'Optimizing FiveM',          image: 'jylli_logo'  },
  'game-opti':        { details: 'Optimizing Games',          image: 'jylli_logo'  },
  'debloater':        { details: 'Removing Bloatware',        image: 'jylli_logo'  },
  'preflight':        { details: 'Running Pre-Flight',        image: 'jylli_logo'  },
  'fixes':            { details: 'Applying Fixes',            image: 'jylli_logo'  },
  'process-manager':  { details: 'Viewing Processes',          image: 'jylli_logo'  },
  'cleaner':          { details: 'Cleaning System',           image: 'jylli_logo'  },
  'scheduler':        { details: 'Scheduling Tweaks',         image: 'jylli_logo'  },
  'changelog':        { details: 'Reading Changelog',         image: 'jylli_logo'  },
  'restore-points':   { details: 'Managing Restore Points',   image: 'jylli_logo'  },
  'startup-manager':  { details: 'Managing Startup Apps',     image: 'jylli_logo'  },
  'ping-test':        { details: 'Testing Network Ping',      image: 'jylli_logo'  },
  'app-optimizer':    { details: 'Optimizing Apps',           image: 'jylli_logo'  },
}

// Cycling phrases shown during long scans
const SCAN_PHRASES  = ['Scanning system…', 'Checking processes…', 'Analysing files…', 'Reading registry…']
const HEALTH_PHRASES = ['Checking system health…', 'Running SFC scan…', 'Checking DISM…', 'Checking disk SMART…']
const OPTI_PHRASES  = ['Auto-Optimizing…', 'Applying tweaks…', 'Writing registry…', 'Configuring system…']
let cycleIndex = 0

function initDiscordRPC() {
  try {
    const RPC = require('discord-rpc')
    discordRPC = new RPC.Client({ transport: 'ipc' })
    discordStartTimestamp = new Date()

    discordRPC.on('ready', () => {
      discordReady = true
      if (discordRetryTimer) { clearInterval(discordRetryTimer); discordRetryTimer = null }
      mainWindow?.webContents.send('discord-rpc-ready')
      updateDiscordPresence()
    })

    discordRPC.on('disconnected', () => {
      discordReady = false
      scheduleDiscordRetry()
    })

    connectDiscordRPC()
  } catch (_) { discordRPC = null }
}

function scheduleDiscordRetry() {
  if (!discordRetryTimer) {
    discordRetryTimer = setInterval(() => {
      if (!discordReady) connectDiscordRPC()
    }, 15000)
  }
}

function connectDiscordRPC() {
  if (!discordRPC) return
  discordRPC.login({ clientId: DISCORD_CLIENT_ID }).catch(() => {
    // Login failed (Discord not running yet) — start retry loop
    scheduleDiscordRetry()
  })
}

function updateDiscordPresence() {
  if (!discordRPC || !discordReady) return

  let details, state, imageKey, imageText, smallKey, smallText

  if (discordActivityOverride) {
    // Active long operation (auto-opti, health scan, etc.) takes top priority
    details   = discordActivityOverride.details
    state     = discordActivityOverride.state
    imageKey  = discordActivityOverride.imageKey  || 'jylli_logo'
    imageText = discordActivityOverride.imageText || 'Jylli Tool'
    smallKey  = 'jylli_logo'
    smallText = 'Jylli Tool'
  } else if (discordPulseGame) {
    // Pulse is active for a specific game
    details   = `⚡ Pulse Active — ${discordPulseGame}`
    state     = discordFpsGained > 0
      ? `+${discordFpsGained} FPS · ${discordTweaksApplied} tweak${discordTweaksApplied !== 1 ? 's' : ''} applied`
      : `${discordTweaksApplied} tweak${discordTweaksApplied !== 1 ? 's' : ''} applied`
    imageKey  = 'jylli_logo'
    imageText = 'Pulse — Live Game Optimization · discord.gg/jylligraffat'
    smallKey  = discordHasFiveM ? 'fivem_icon' : 'jylli_small'
    smallText = discordHasFiveM ? 'FiveM detected' : 'Gaming Mode'
  } else if (discordActiveGame) {
    // Game watcher detected a running game
    details   = `🎮 In-Game — ${discordActiveGame}`
    state     = discordFpsGained > 0
      ? `+${discordFpsGained} FPS gained from tweaks`
      : `${discordTweaksApplied} tweak${discordTweaksApplied !== 1 ? 's' : ''} applied`
    imageKey  = 'jylli_logo'
    imageText = 'Jylli Tool — Game Watcher Active · discord.gg/jylligraffat'
    smallKey  = discordHasFiveM ? 'fivem_icon' : 'jylli_small'
    smallText = discordHasFiveM ? 'FiveM detected' : 'Jylli Tool'
  } else {
    // Idle — show current page
    const p   = PAGE_PRESENCE[discordCurrentPage] || { details: 'Using Jylli Tool', image: 'jylli_logo' }
    details   = p.details
    state     = discordFpsGained > 0
      ? `+${discordFpsGained} FPS gained · ${discordTweaksApplied} tweak${discordTweaksApplied !== 1 ? 's' : ''} applied`
      : discordTweaksApplied > 0
        ? `${discordTweaksApplied} tweak${discordTweaksApplied !== 1 ? 's' : ''} applied`
        : 'discord.gg/jylligraffat'
    imageKey  = p.image
    imageText = 'Jylli Tool — Windows Optimizer · discord.gg/jylligraffat'
    smallKey  = discordHasFiveM ? 'fivem_icon' : 'jylli_small'
    smallText = discordHasFiveM ? 'FiveM detected' : 'Jylli Tool'
  }

  // Elapsed session time shown as "elapsed" timer (counts up from session start)
  discordRPC.setActivity({
    details,
    state,
    startTimestamp: discordStartTimestamp,
    largeImageKey:  imageKey,
    largeImageText: imageText,
    smallImageKey:  smallKey,
    smallImageText: smallText,
    buttons: [
      { label: 'Join Discord',        url: DISCORD_INVITE   },
      { label: 'Download Jylli Tool', url: DISCORD_DOWNLOAD },
    ],
    instance: false,
  }).catch(() => {})
}

// Start a cycling "active operation" override (clears after calling clearDiscordOverride)
function setDiscordOverride(phrases, stepLabel) {
  cycleIndex = 0
  if (discordOverrideCycleTimer) { clearInterval(discordOverrideCycleTimer); discordOverrideCycleTimer = null }
  discordActivityOverride = { details: phrases[0], state: stepLabel || '', imageKey: 'jylli_logo' }
  updateDiscordPresence()
  discordOverrideCycleTimer = setInterval(() => {
    cycleIndex = (cycleIndex + 1) % phrases.length
    if (discordActivityOverride) {
      discordActivityOverride.details = phrases[cycleIndex]
      if (stepLabel !== undefined) discordActivityOverride.state = stepLabel
    }
    updateDiscordPresence()
  }, 4000)
}

function updateDiscordOverrideState(stepLabel) {
  if (discordActivityOverride) {
    discordActivityOverride.state = stepLabel
    updateDiscordPresence()
  }
}

function clearDiscordOverride() {
  if (discordOverrideCycleTimer) { clearInterval(discordOverrideCycleTimer); discordOverrideCycleTimer = null }
  discordActivityOverride = null
  updateDiscordPresence()
}

async function destroyDiscordRPC() {
  clearDiscordOverride()
  if (discordRetryTimer) { clearInterval(discordRetryTimer); discordRetryTimer = null }
  if (discordRPC) {
    try {
      // Wait for clearActivity to complete before destroying — otherwise status lingers
      if (discordReady) await discordRPC.clearActivity().catch(() => {})
      discordRPC.destroy()
    } catch (_) {}
    discordRPC = null
    discordReady = false
  }
}

// ─── Analytics Webhook ───────────────────────────────────────────────────────
const WEBHOOK_URL    = 'https://discord.com/api/webhooks/1504473739599941682/FrOgbUEKp9_1jwSKX2o4gbBAS2FSFbgEnVNVte5s9DDoBS8J7_1aC-7JzWG7K1ZZ66BZ'
const ANALYTICS_PATH = path.join(app.getPath('userData'), 'jt_analytics.json')
const APP_VERSION    = app.getVersion()
ipcMain.handle('get-app-version', () => APP_VERSION)

// Brand colors (matches app theme)
const COLOR_PINK    = 0xFF2E63  // primary accent — launches, returning users
const COLOR_GREEN   = 0x2ECC71  // new installs
const COLOR_ERROR   = 0xFF2E63  // crashes
const COLOR_WARNING = 0xFF6B35  // non-fatal errors
const COLOR_BLUE    = 0x3498DB  // auto-opti complete
const COLOR_PURPLE  = 0x9B59B6  // session end / uninstall

// Session start time — used to calculate duration on quit
let sessionStartTime = Date.now()

function loadAnalytics() {
  try {
    if (fs.existsSync(ANALYTICS_PATH)) return JSON.parse(fs.readFileSync(ANALYTICS_PATH, 'utf8'))
  } catch {}
  return {}
}

function saveAnalytics(data) {
  try { fs.writeFileSync(ANALYTICS_PATH, JSON.stringify(data, null, 2)) } catch {}
}

function sendWebhook(payload) {
  try {
    const https = require('https')
    const body = JSON.stringify(payload)
    const url = new URL(WEBHOOK_URL)
    const req = https.request({
      hostname: url.hostname, path: url.pathname + url.search,
      method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) }
    })
    req.on('error', () => {})
    req.write(body)
    req.end()
  } catch {}
}

// ── Shared helper: build system info fields ───────────────────────────────────
function getSystemFields() {
  try {
    const osMod  = require('os')
    const rel    = osMod.release()
    const build  = parseInt(rel.split('.')[2] || '0')
    const winVer = build >= 22000 ? `Windows 11 (Build ${build})` : `Windows 10 (Build ${build})`
    const ram    = `${Math.round(osMod.totalmem() / 1073741824)} GB`
    const cpus   = osMod.cpus()
    const cpu    = cpus?.length ? cpus[0].model.replace(/\s+/g, ' ').trim() : 'Unknown'
    const hasFiveM = fs.existsSync(path.join(process.env.LOCALAPPDATA || '', 'FiveM', 'FiveM.app'))
    // GPU — cached from last detectSystemInfo run if available, otherwise skip
    const gpu    = getSystemFields._cachedGpu || null
    const isLaptop = getSystemFields._cachedIsLaptop || false
    const isWifi   = getSystemFields._cachedIsWifi   || false
    return { winVer, ram, cpu, hasFiveM, gpu, isLaptop, isWifi }
  } catch { return { winVer: 'Unknown', ram: '?', cpu: 'Unknown', hasFiveM: false, gpu: null, isLaptop: false, isWifi: false } }
}
// Populated after detectSystemInfo completes
getSystemFields._cachedGpu      = null
getSystemFields._cachedIsLaptop = false
getSystemFields._cachedIsWifi   = false

function webhookLaunch() {
  try {
    const analytics = loadAnalytics()
    const now    = new Date()
    const nowIso = now.toISOString()
    const nowMs  = now.getTime()

    // First-ever run on this machine
    const isNewUser = !analytics.userId
    if (isNewUser) {
      analytics.userId    = require('crypto').randomBytes(8).toString('hex')
      analytics.firstSeen = nowIso
      analytics.sessions  = 0
      analytics.totalTweaksApplied = 0
    }

    analytics.sessions      = (analytics.sessions || 0) + 1
    analytics.version       = APP_VERSION
    const lastSeenIso       = analytics.lastSeen || null
    analytics.lastSeen      = nowIso

    // Track session dates for activity windows (keep 90 days)
    analytics.sessionDates = (analytics.sessionDates || [])
      .filter(d => nowMs - new Date(d).getTime() < 90 * 86400000)
    analytics.sessionDates.push(nowIso)

    // Track versions seen
    analytics.versionHistory = analytics.versionHistory || []
    if (!analytics.versionHistory.includes(APP_VERSION)) analytics.versionHistory.push(APP_VERSION)

    saveAnalytics(analytics)

    const dau7  = analytics.sessionDates.filter(d => nowMs - new Date(d).getTime() < 7  * 86400000).length
    const dau30 = analytics.sessionDates.filter(d => nowMs - new Date(d).getTime() < 30 * 86400000).length
    const { winVer, ram, cpu, hasFiveM, gpu, isLaptop, isWifi } = getSystemFields()

    const ts          = Math.floor(nowMs / 1000)
    const firstSeenTs = Math.floor(new Date(analytics.firstSeen).getTime() / 1000)
    const daysSinceLast = lastSeenIso
      ? Math.floor((nowMs - new Date(lastSeenIso).getTime()) / 86400000) : null
    const returnLabel = daysSinceLast === null ? '—'
      : daysSinceLast === 0 ? 'Today'
      : daysSinceLast === 1 ? '1 day ago'
      : `${daysSinceLast} days ago`

    const deviceType = isLaptop ? '💻 Laptop' : '🖥️ Desktop'
    const netType    = isWifi   ? '📶 Wi-Fi'  : '🔌 Ethernet'

    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: isNewUser ? COLOR_GREEN : COLOR_PINK,
        title: isNewUser ? '🎉  New User' : '▶  User Session',
        description: isNewUser
          ? `First ever launch on this machine. **Welcome to Jylli Tool v${APP_VERSION}!**`
          : `Returning user — last seen **${returnLabel}** · session **#${analytics.sessions}**`,
        fields: [
          { name: '🪪  User ID',        value: `\`${analytics.userId}\``,                                  inline: true },
          { name: '📦  Version',        value: `**v${APP_VERSION}**`,                                      inline: true },
          { name: '📅  Time',           value: `<t:${ts}:f>`,                                              inline: true },
          { name: '🖥️  OS',             value: winVer,                                                     inline: true },
          { name: '⚙️  CPU',            value: `\`${cpu.length > 28 ? cpu.slice(0, 28) + '…' : cpu}\``,   inline: true },
          { name: '🧠  RAM',            value: ram,                                                        inline: true },
          ...(gpu ? [{ name: '🎨  GPU', value: `\`${gpu.length > 28 ? gpu.slice(0, 28) + '…' : gpu}\``,   inline: true }] : []),
          { name: '🎮  FiveM',          value: hasFiveM ? '✅ Installed' : '❌ Not found',                 inline: true },
          { name: '🖱️  Device',         value: deviceType,                                                 inline: true },
          { name: '🌐  Network',        value: netType,                                                    inline: true },
          { name: '📊  Active (7d)',    value: `**${dau7}** session${dau7  !== 1 ? 's' : ''}`,             inline: true },
          { name: '📈  Active (30d)',   value: `**${dau30}** session${dau30 !== 1 ? 's' : ''}`,            inline: true },
          { name: '📆  First Seen',     value: `<t:${firstSeenTs}:D>`,                                     inline: true },
          { name: '🔧  Tweaks Ever',    value: `**${analytics.totalTweaksApplied || 0}**`,                 inline: true },
          { name: '🔢  Total Sessions', value: `**${analytics.sessions}**`,                                inline: true },
          { name: '📚  Versions Used',  value: analytics.versionHistory?.join(', ') || `v${APP_VERSION}`, inline: true },
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Analytics` },
        timestamp: nowIso,
      }]
    })
  } catch {}
}

function webhookSessionEnd(durationMs, pageVisits, tweaksThisSession, pulseUses) {
  try {
    const analytics = loadAnalytics()
    if (!analytics.userId) return

    const mins = Math.floor(durationMs / 60000)
    const secs = Math.floor((durationMs % 60000) / 1000)
    const durationLabel = mins > 0 ? `${mins}m ${secs}s` : `${secs}s`

    // Top 3 most visited pages this session
    const sortedPages = Object.entries(pageVisits || {}).sort((a, b) => b[1] - a[1])
    const topPagesLabel = sortedPages.length
      ? sortedPages.slice(0, 3).map(([p, n]) => `**${p}** ×${n}`).join(', ')
      : '—'

    // Update lifetime tweaks count
    analytics.totalTweaksApplied = (analytics.totalTweaksApplied || 0) + (tweaksThisSession || 0)
    analytics.totalPulseUses = (analytics.totalPulseUses || 0) + (pulseUses || 0)
    saveAnalytics(analytics)

    const ts = Math.floor(Date.now() / 1000)
    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: COLOR_PURPLE,
        title: '🔴  Session Ended',
        description: `User closed Jylli Tool after **${durationLabel}**.`,
        fields: [
          { name: '🪪  User ID',         value: `\`${analytics.userId}\``,             inline: true },
          { name: '⏱️  Duration',         value: `**${durationLabel}**`,                inline: true },
          { name: '📅  Time',            value: `<t:${ts}:f>`,                          inline: true },
          { name: '📄  Top Pages',       value: topPagesLabel,                          inline: false },
          { name: '🔧  Tweaks Applied',  value: `**${tweaksThisSession || 0}**`,        inline: true },
          { name: '🔧  Lifetime Tweaks', value: `**${analytics.totalTweaksApplied}**`,  inline: true },
          { name: '⚡  Pulse Uses',      value: `**${pulseUses || 0}** this session · **${analytics.totalPulseUses}** lifetime`, inline: false },
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Analytics` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

function webhookPulseActivated(gameName, optimizations) {
  try {
    const analytics = loadAnalytics()
    const userId = analytics.userId || 'unknown'
    const ts = Math.floor(Date.now() / 1000)
    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: COLOR_PINK,
        title: '⚡  Pulse Activated',
        description: `User activated Pulse for **${gameName}**.`,
        fields: [
          { name: '🪪  User ID',         value: `\`${userId}\``,              inline: true },
          { name: '🎮  Game',            value: `**${gameName}**`,             inline: true },
          { name: '📅  Time',            value: `<t:${ts}:f>`,                 inline: true },
          { name: '🔩  Optimizations',   value: optimizations.map(o => `• ${o}`).join('\n'), inline: false },
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Analytics` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

function webhookAutoOpti(tweakCount, fpsEstimate, includedFiveM) {
  try {
    const analytics = loadAnalytics()
    const userId = analytics.userId || 'unknown'
    analytics.autoOptiRuns = (analytics.autoOptiRuns || 0) + 1
    analytics.totalTweaksApplied = (analytics.totalTweaksApplied || 0) + tweakCount
    saveAnalytics(analytics)

    const ts = Math.floor(Date.now() / 1000)
    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: COLOR_BLUE,
        title: '⚡  Auto-Optimize Completed',
        description: `User ran Auto-Optimize and applied **${tweakCount} tweaks** — estimated **~+${fpsEstimate} FPS** gained.`,
        fields: [
          { name: '🪪  User ID',       value: `\`${userId}\``,                           inline: true },
          { name: '🔧  Tweaks Applied',value: `**${tweakCount}**`,                        inline: true },
          { name: '📈  FPS Estimate',  value: `**~+${fpsEstimate} FPS**`,                 inline: true },
          { name: '🎮  FiveM Tweaks',  value: includedFiveM ? '✅ Included' : '❌ Skipped', inline: true },
          { name: '🔁  Run #',         value: `**${analytics.autoOptiRuns}**`,            inline: true },
          { name: '📅  Time',          value: `<t:${ts}:f>`,                              inline: true },
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Analytics` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

function webhookHealthCheck(results) {
  try {
    const analytics = loadAnalytics()
    const userId = analytics.userId || 'unknown'

    const sfcLabel  = results.sfc  === 'clean'    ? '✅ Clean'
      : results.sfc  === 'corrupt'  ? '❌ Corrupt files found'
      : results.sfc  === 'pending'  ? '⏳ Pending reboot'
      : '❓ Unknown'
    const dismLabel = results.dism === 'healthy'   ? '✅ Healthy'
      : results.dism === 'repairable' ? '⚠️ Repairable'
      : results.dism === 'corrupted'  ? '❌ Corrupted'
      : '❓ Unknown'
    const diskLabel = results.disk === 'healthy'   ? '✅ Healthy'
      : results.disk === 'warning'    ? '⚠️ Warning'
      : results.disk === 'failing'    ? '❌ Failing'
      : '❓ Unknown'

    const hasIssues = !results.ok
    const ts = Math.floor(Date.now() / 1000)

    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: hasIssues ? COLOR_WARNING : COLOR_GREEN,
        title: hasIssues ? '⚠️  Windows Health — Issues Found' : '✅  Windows Health — All Clear',
        description: hasIssues
          ? `Health check found **${results.issues?.length || 1} issue(s)** on this system.`
          : `System passed all health checks.`,
        fields: [
          { name: '🪪  User ID',  value: `\`${userId}\``, inline: true },
          { name: '📅  Time',     value: `<t:${ts}:f>`,   inline: true },
          { name: '​',       value: '​',         inline: true },
          { name: '🛡️  SFC',      value: sfcLabel,         inline: true },
          { name: '🏥  DISM',     value: dismLabel,        inline: true },
          { name: '💾  Disk',     value: diskLabel,        inline: true },
          ...(hasIssues ? [{ name: '📋  Issues', value: results.issues.map(i => `• ${i}`).join('\n').slice(0, 900) }] : []),
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Analytics` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

function webhookError(type, err) {
  try {
    const analytics = loadAnalytics()
    const userId = analytics.userId || 'unknown'
    const isCrash = type === 'crash'
    const msg = err?.message || String(err)
    const stack = (err?.stack || '')
      .split('\n')
      .slice(0, 5)
      .join('\n')
      .trim()

    const ts = Math.floor(Date.now() / 1000)

    sendWebhook({
      username: 'Jylli Tool',
      embeds: [{
        color: isCrash ? COLOR_ERROR : COLOR_WARNING,
        title: isCrash ? '💥  App Crash' : '⚠️  Unhandled Error',
        description: isCrash
          ? 'Jylli Tool encountered a fatal error and crashed.'
          : 'An unhandled error occurred but the app kept running.',
        fields: [
          { name: '🪪  User',   value: `\`${userId}\``,   inline: true },
          { name: '📅  Time',   value: `<t:${ts}:f>`,     inline: true },
          { name: '🔖  Type',   value: `\`${type}\``,     inline: true },
          { name: '❌  Error',  value: `\`\`\`\n${msg.slice(0, 800)}\n\`\`\`` },
          ...(stack ? [{ name: '📋  Stack Trace', value: `\`\`\`\n${stack.slice(0, 800)}\n\`\`\`` }] : []),
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Error Report` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

function webhookBugReport({ title, description, steps, page }) {
  try {
    const analytics = loadAnalytics()
    const userId = analytics.userId || 'unknown'
    const { winVer, cpu, ram, gpu, isLaptop, isWifi } = getSystemFields()
    const ts = Math.floor(Date.now() / 1000)
    const deviceType = isLaptop ? '💻 Laptop' : '🖥️ Desktop'
    const netType    = isWifi   ? '📶 Wi-Fi'  : '🔌 Ethernet'
    sendWebhook({
      username: 'Jylli Tool — Bug Report',
      embeds: [{
        color: 0xE74C3C,
        title: `🐛  Bug Report — ${title.slice(0, 80)}`,
        description: description.slice(0, 1800),
        fields: [
          { name: '🪪  User ID',      value: `\`${userId}\``,                                                inline: true },
          { name: '📦  Version',      value: `**v${APP_VERSION}**`,                                         inline: true },
          { name: '📅  Time',         value: `<t:${ts}:f>`,                                                 inline: true },
          { name: '📄  Page',         value: page || '—',                                                   inline: true },
          { name: '🖱️  Device',        value: deviceType,                                                   inline: true },
          { name: '🌐  Network',      value: netType,                                                       inline: true },
          { name: '🖥️  OS',           value: winVer,                                                        inline: true },
          { name: '⚙️  CPU',          value: `\`${(cpu||'?').slice(0, 28)}\``,                              inline: true },
          ...(gpu ? [{ name: '🎨  GPU', value: `\`${gpu.slice(0, 28)}\``,                                   inline: true }] : []),
          ...(steps ? [{ name: '📋  Steps to Reproduce', value: steps.slice(0, 900) }] : []),
        ],
        footer: { text: `Jylli Tool  ·  v${APP_VERSION}  ·  Bug Report` },
        timestamp: new Date().toISOString(),
      }]
    })
  } catch {}
}

ipcMain.handle('submit-bug-report', (_, data) => {
  webhookBugReport(data)
  return { ok: true }
})

// Wire up global error handlers
process.on('uncaughtException',  (err) => { webhookError('crash', err) })
process.on('unhandledRejection', (err) => { webhookError('unhandledRejection', err) })

// ─── Window ───────────────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 1060,
    minHeight: 660,
    frame: false,
    transparent: false,
    backgroundColor: '#0d0d0d',
    show: false,
    icon: path.join(__dirname, 'assets', 'icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })

  mainWindow.loadFile('index.html')

  mainWindow.once('ready-to-show', () => {
    mainWindow.show()
  })

  // Minimize → hide to tray instead of taskbar minimize
  mainWindow.on('minimize', (e) => {
    e.preventDefault()
    mainWindow.hide()
    updateTrayMenu()
    if (discordReady) discordRPC.clearActivity().catch(() => {})
  })

  // X button quits the app fully
  mainWindow.on('close', () => {
    app.isQuitting = true
    app.quit()
  })

  mainWindow.on('closed', () => { mainWindow = null })
}

function updateTrayMenu() {
  if (!tray) return
  const pulseOn = !!pulseActivePreset
  const autoPulseOn = !!autoPulseTriggeredPreset
  const shown = mainWindow && mainWindow.isVisible()

  const menu = Menu.buildFromTemplate([
    {
      label: 'Jylli Tool',
      enabled: false,
      icon: nativeImage.createFromPath(path.join(__dirname, 'assets', 'icon.ico')).resize({ width: 16, height: 16 }),
    },
    { type: 'separator' },
    {
      label: pulseOn
        ? `⚡ Pulse ACTIVE${autoPulseOn ? ' (Auto)' : ''} — ${pulseActivePreset || ''}`
        : '  Pulse: inactive',
      enabled: false,
    },
    {
      label: pulseOn ? '⏹  Stop Pulse' : '⚡ Start Pulse (auto-detect)',
      click: async () => {
        const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
        if (pulseOn) {
          autoPulseTriggeredPreset = null
          await deactivatePulse(send)
          mainWindow?.webContents.send('tray-pulse-changed', { active: false })
        } else {
          // Detect running game first, fall back to last used preset
          let presetId = null
          try {
            const r = await runPS('Get-Process | Select-Object -ExpandProperty Name', 5000)
            const lines = r.out.toLowerCase().split('\n').map(l => l.trim())
            for (const [id, p] of Object.entries(PULSE_PRESETS)) {
              const pattern = p.exe.toLowerCase()
              const match = pattern.includes('_gtaprocess')
                ? lines.some(l => l.endsWith('_gtaprocess'))
                : pattern.endsWith('*')
                  ? lines.some(l => l.startsWith(pattern.slice(0, -1)))
                  : lines.some(l => l === pattern)
              if (match) { presetId = id; break }
            }
          } catch {}
          presetId = presetId || lastUsedPulsePreset || Object.keys(PULSE_PRESETS)[0]
          await activatePulse(presetId, false, send)
          mainWindow?.webContents.send('tray-pulse-changed', { active: true, presetId })
        }
        updateTrayMenu()
      },
    },
    { type: 'separator' },
    {
      label: shown ? 'Hide window' : 'Show window',
      click: () => {
        if (shown) { mainWindow.hide(); if (discordReady) discordRPC.clearActivity().catch(() => {}) }
        else { mainWindow.show(); mainWindow.focus(); updateDiscordPresence() }
        updateTrayMenu()
      },
    },
    {
      label: 'Quit Jylli Tool',
      click: () => {
        app.isQuitting = true
        app.quit()
      },
    },
  ])

  tray.setContextMenu(menu)
  tray.setToolTip(pulseOn ? `Jylli Tool — Pulse active (${pulseActivePreset})` : 'Jylli Tool')
}

function createTray() {
  const iconPath = path.join(__dirname, 'assets', 'icon.ico')
  tray = new Tray(iconPath)
  tray.setToolTip('Jylli Tool')

  // Single click restores the window
  tray.on('click', () => {
    if (mainWindow) {
      if (mainWindow.isVisible()) { mainWindow.focus() }
      else { mainWindow.show(); mainWindow.focus(); updateDiscordPresence() }
    }
    updateTrayMenu()
  })

  updateTrayMenu()
}

// ─── Auto-elevation ───────────────────────────────────────────────────────────
// npm start doesn't pass the admin token to Electron child processes.
// Check elevation on ready and relaunch via UAC prompt if not admin.
if (IS_WIN) {
  app.whenReady().then(async () => {
    const { execSync } = require('child_process')
    try {
      const r = await new Promise((resolve) => {
        execFile('powershell.exe', [
          '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden',
          '-ExecutionPolicy', 'Bypass', '-Command',
          '([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)'
        ], { windowsHide: true, encoding: 'utf8' }, (err, stdout) => resolve((stdout || '').trim().toLowerCase()))
      })
      if (r !== 'true') {
        const appPath = app.getAppPath()
        const exePath = process.execPath
        execSync(
          `powershell -WindowStyle Hidden -Command "Start-Process '${exePath}' -ArgumentList '--inspect=0','${appPath}' -Verb RunAs"`,
          { windowsHide: true }
        )
        app.quit()
      }
    } catch (_) {
      // UAC was declined or failed — continue without admin
    }
  })
}

// Session-level tracking (populated by renderer via IPC)
let sessionPageVisits   = {}
let sessionTweaksCount  = 0
let sessionPulseUses    = 0

app.whenReady().then(() => { createWindow(); createTray(); initDiscordRPC(); webhookLaunch() })
app.on('window-all-closed', () => { /* stay alive in tray */ })
app.on('will-quit', (e) => {
  e.preventDefault()
  const duration = Date.now() - sessionStartTime
  if (duration > 10000) webhookSessionEnd(duration, sessionPageVisits, sessionTweaksCount, sessionPulseUses)
  // Await clearActivity so Discord status is actually removed before the process exits
  destroyDiscordRPC().finally(() => app.exit(0))
})

// ── Session tracking IPC ─────────────────────────────────────────────────────
ipcMain.handle('analytics-page-visit', (_, page) => {
  sessionPageVisits[page] = (sessionPageVisits[page] || 0) + 1
})
ipcMain.handle('analytics-tweak-applied', () => {
  sessionTweaksCount++
})

// ─── PowerShell runner ────────────────────────────────────────────────────────
function runPS(command, timeoutMs = 30000) {
  return new Promise((resolve) => {
    const args = [
      '-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden',
      '-ExecutionPolicy', 'Bypass', '-Command', command
    ]
    const proc = execFile('powershell.exe', args, {
      timeout: timeoutMs,
      windowsHide: true,
      encoding: 'utf8'
    }, (err, stdout, stderr) => {
      resolve({
        ok: !err || err.code === 0,
        out: (stdout || '').trim(),
        err: (stderr || err?.message || '').trim(),
        code: err?.code ?? 0
      })
    })
  })
}

// Run a simple cmd/system binary (sc, bcdedit, netsh, etc.) fully hidden
function runCmd(cmd) {
  return new Promise((resolve) => {
    exec(cmd, {
      windowsHide: true,
      timeout: 30000
    }, (err, stdout, stderr) => {
      resolve({
        ok: !err || err.code === 0,
        out: (stdout || '').trim(),
        err: (stderr || '').trim(),
        code: err?.code ?? 0
      })
    })
  })
}

// Registry helpers via reg.exe (no PowerShell overhead for simple reg writes)
function regAdd(hive, path_, name, type, value) {
  return runCmd(`reg add "${hive}\\${path_}" /v "${name}" /t ${type} /d "${value}" /f`)
}
function regDelete(hive, path_, name) {
  return runCmd(`reg delete "${hive}\\${path_}" /v "${name}" /f`)
}

// ─── Admin detection ─────────────────────────────────────────────────────────
async function checkAdmin() {
  if (!IS_WIN) return false
  const r = await runPS('([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)')
  return r.out.toLowerCase().trim() === 'true'
}

// ─── Settings persistence ────────────────────────────────────────────────────
function loadSettings() {
  try {
    if (fs.existsSync(SETTINGS_PATH)) return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'))
  } catch {}
  return {}
}
function saveSettings(data) {
  try { fs.writeFileSync(SETTINGS_PATH, JSON.stringify(data, null, 2)) } catch {}
}

// ─── System Info Detection ────────────────────────────────────────────────────
function emitSpecProgress(label, detail = '') {
  mainWindow?.webContents.send('spec-progress', { label, detail })
}

async function detectSystemInfo() {
  const info = {
    cpu: 'Detecting…', cpuCores: os.cpus().length, cpuSpeed: 0,
    ram: 0, ramGB: 0, ramSpeed: 0,
    gpu: 'Detecting…', vramMB: 0,
    os: os.version?.() ?? os.release(),
    osCaption: '',
    isAdmin: false,
    isLaptop: false,
    isWifi: false,
    wifiAdapter: '',
    nvme: false,
    diskUsedGB: 0,
    diskTotalGB: 0,
    diskType: 'SSD/HDD',
    arch: os.arch()
  }

  // CPU — get real max speed from WMI, not the throttled current speed from os.cpus()
  const cpus = os.cpus()
  if (cpus.length > 0) {
    info.cpu = cpus[0].model.replace(/\s+/g, ' ').trim()
    info.cpuCores = cpus.length
    // os.cpus().speed is the CURRENT (idle-throttled) speed — often 800MHz on boosting CPUs.
    // Get the real MaxClockSpeed from WMI instead.
    info.cpuSpeed = cpus[0].speed // fallback, overwritten below if WMI succeeds
  }

  // RAM
  info.ram = os.totalmem()
  info.ramGB = Math.floor(info.ram / (1024 ** 3))
  if (IS_WIN) {
    const ramSpeedR = await runPS('(Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property ConfiguredClockSpeed -Maximum).Maximum')
    const spd = parseInt(ramSpeedR.out.trim())
    if (!isNaN(spd) && spd > 0) info.ramSpeed = spd
  }

  if (IS_WIN) {
    // ── CPU real max speed from WMI ──────────────────────────────────────────
    emitSpecProgress('Detecting CPU…', info.cpu || '')
    const cpuSpeedR = await runPS('(Get-CimInstance Win32_Processor | Select-Object -First 1).MaxClockSpeed')
    const maxSpeed = parseInt(cpuSpeedR.out.trim())
    if (!isNaN(maxSpeed) && maxSpeed > 0) info.cpuSpeed = maxSpeed
    emitSpecProgress('CPU detected', `${info.cpu} · ${(info.cpuSpeed/1000).toFixed(1)} GHz · ${info.cpuCores} cores`)
    await new Promise(r => setTimeout(r, 50))

    // ── GPU Name ─────────────────────────────────────────────────────────────
    // Get GPU name from CIM (most reliable for name only)
    emitSpecProgress('Detecting GPU…', '')
    const gpuNameR = await runPS(
      'Get-CimInstance Win32_VideoController | Where-Object {$_.Name -notlike "*Basic*" -and $_.Name -notlike "*Remote*" -and $_.Name -notlike "*Virtual*"} | Sort-Object AdapterRAM -Descending | Select-Object -First 1 -ExpandProperty Name'
    )
    if (gpuNameR.ok && gpuNameR.out.trim()) info.gpu = gpuNameR.out.trim()
    emitSpecProgress('GPU detected', info.gpu)
    await new Promise(r => setTimeout(r, 50))

    // ── VRAM — Complete rewrite using DXGI C# + dxdiag + HWiNFO ──────────────
    //
    // Root cause of previous failures:
    //   • Win32_VideoController.AdapterRAM = 32-bit DWORD, caps at 4GB
    //   • HardwareInformation.MemorySize is stored as REG_BINARY on most drivers
    //     so parseInt() returns NaN — the value is never read correctly
    //
    // New approach (4-method waterfall, stops at first success > 512MB):
    //   1. DXGI via C# P/Invoke — exact same API Task Manager uses
    //   2. dxdiag /t XML output — parses "Display Memory" field
    //   3. HWiNFO64 shared memory segment — if HWiNFO is running
    //   4. REG_BINARY manual parse — reads the raw bytes correctly

    emitSpecProgress('Detecting VRAM…', 'Querying DXGI…')
    // ── Method 1: DXGI IDXGIAdapter1::GetDesc1 via inline C# ─────────────────
    // This is what Task Manager / GPU-Z / DirectX Diagnostic use internally.
    // DedicatedVideoMemory is a SIZE_T (64-bit on x64) — no overflow.
    const dxgiResult = await runPS(`
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DXGIHelper {
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct DXGI_ADAPTER_DESC1 {
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)] public string Description;
    public uint VendorId, DeviceId, SubSysId, Revision;
    public IntPtr DedicatedVideoMemory, DedicatedSystemMemory, SharedSystemMemory;
    public long AdapterLuid;
    public uint Flags;
  }
  [DllImport("dxgi.dll")] static extern int CreateDXGIFactory1(ref Guid riid, out IntPtr ppFactory);
  static readonly Guid IID_IDXGIFactory1 = new Guid("770aae78-f26f-4dba-a829-253c83d1b387");
  [ComImport, Guid("770aae78-f26f-4dba-a829-253c83d1b387"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IDXGIFactory1 { void stub0(); void stub1(); void stub2(); int EnumAdapters1(uint idx, out IDXGIAdapter1 adapter); }
  [ComImport, Guid("29038f61-3839-4626-91fd-086879011a05"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IDXGIAdapter1 { void stub0(); void stub1(); void stub2(); void stub3(); void stub4(); int GetDesc1(out DXGI_ADAPTER_DESC1 desc); }
  public static long GetVRAMBytes() {
    try {
      Guid iid = IID_IDXGIFactory1; IntPtr pFac;
      if (CreateDXGIFactory1(ref iid, out pFac) != 0) return 0;
      var fac = (IDXGIFactory1)Marshal.GetObjectForIUnknown(pFac);
      IDXGIAdapter1 adapter; long best = 0;
      for (uint i = 0; fac.EnumAdapters1(i, out adapter) == 0; i++) {
        DXGI_ADAPTER_DESC1 desc;
        adapter.GetDesc1(out desc);
        long vram = desc.DedicatedVideoMemory.ToInt64();
        if (vram > best && desc.VendorId != 0x1414) best = vram;
        Marshal.ReleaseComObject(adapter);
      }
      Marshal.ReleaseComObject(fac); Marshal.Release(pFac);
      return best;
    } catch { return 0; }
  }
}
'@ -ErrorAction SilentlyContinue
try { $v = [DXGIHelper]::GetVRAMBytes(); Write-Output $v } catch { Write-Output 0 }
    `, 20000)

    const dxgiBytes = parseInt(dxgiResult.out.trim())
    if (!isNaN(dxgiBytes) && dxgiBytes > 536870912) {
      info.vramMB = Math.floor(dxgiBytes / (1024 ** 2))
    }

    // ── Method 2: dxdiag /t — parse XML "Approx. Total Memory" or "Display Memory" ──
    if (info.vramMB < 512) {
      emitSpecProgress('Detecting VRAM…', 'Running dxdiag…')
      const dxdiagResult = await runPS(`
        $tmp = [System.IO.Path]::GetTempFileName() + '.txt'
        try {
          Start-Process dxdiag -ArgumentList '/t',$tmp -Wait -WindowStyle Hidden -ErrorAction Stop
          if (Test-Path $tmp) {
            $txt = Get-Content $tmp -Raw
            $m = [regex]::Match($txt, 'Dedicated Memory:\\s*(\\d+[\\.,]?\\d*)\\s*MB')
            if (-not $m.Success) { $m = [regex]::Match($txt, 'Display Memory:\\s*(\\d+[\\.,]?\\d*)\\s*MB') }
            if (-not $m.Success) { $m = [regex]::Match($txt, 'Approx.*Memory:\\s*(\\d+[\\.,]?\\d*)\\s*MB') }
            if ($m.Success) { Write-Output ($m.Groups[1].Value -replace ',','') }
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
          }
        } catch {}
      `, 30000)
      const dxMB = parseInt(dxdiagResult.out.trim())
      if (!isNaN(dxMB) && dxMB > 512) info.vramMB = dxMB
    }

    if (info.vramMB >= 512) {
      emitSpecProgress('VRAM detected', `${info.vramMB >= 1024 ? (info.vramMB/1024).toFixed(0)+'GB' : info.vramMB+'MB'} VRAM`)
    }

    // ── Method 3: HWiNFO64 shared memory (if HWiNFO is running) ──────────────
    // HWiNFO exposes a named shared memory segment "Global\\HWiNFO_SENS_SM2"
    // We check if the process is running first to avoid timeout
    if (info.vramMB < 512) {
      const hwInfoRunning = await runPS(
        'if (Get-Process -Name "HWiNFO64","HWiNFO32" -ErrorAction SilentlyContinue) { Write-Output "yes" } else { Write-Output "no" }'
      )
      if (hwInfoRunning.out.trim() === 'yes') {
        const hwInfoResult = await runPS(`
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class HWInfo {
  [DllImport("kernel32")] static extern IntPtr OpenFileMapping(uint access, bool inherit, string name);
  [DllImport("kernel32")] static extern IntPtr MapViewOfFile(IntPtr h, uint access, uint hi, uint lo, UIntPtr size);
  [DllImport("kernel32")] static extern bool UnmapViewOfFile(IntPtr p);
  [DllImport("kernel32")] static extern bool CloseHandle(IntPtr h);
  public static long GetVRAM() {
    IntPtr h = OpenFileMapping(4, false, "Global\\\\HWiNFO_SENS_SM2");
    if (h == IntPtr.Zero) return 0;
    IntPtr p = MapViewOfFile(h, 4, 0, 0, UIntPtr.Zero);
    if (p == IntPtr.Zero) { CloseHandle(h); return 0; }
    try {
      // HWiNFO shared memory: skip header (header size at offset 0), iterate sensors
      int hdrSize = Marshal.ReadInt32(p, 0);
      int sensorSize = Marshal.ReadInt32(p, 4);
      int sensorCount = Marshal.ReadInt32(p, 8);
      int entrySize = Marshal.ReadInt32(p, 12);
      int entryCount = Marshal.ReadInt32(p, 16);
      int entryOff = Marshal.ReadInt32(p, 20);
      for (int i = 0; i < entryCount; i++) {
        IntPtr ep = new IntPtr(p.ToInt64() + entryOff + i * entrySize);
        string name = Marshal.PtrToStringAnsi(new IntPtr(ep.ToInt64() + 0));
        if (name != null && name.ToLower().Contains("gpu") && name.ToLower().Contains("memory")) {
          // value is a double at offset 216
          double val = BitConverter.Int64BitsToDouble(Marshal.ReadInt64(ep, 216));
          if (val > 512) return (long)val;
        }
      }
      return 0;
    } finally { UnmapViewOfFile(p); CloseHandle(h); }
  }
}
'@ -ErrorAction SilentlyContinue
try { $v = [HWInfo]::GetVRAM(); Write-Output $v } catch { Write-Output 0 }
        `, 10000)
        const hwVRAM = parseInt(hwInfoResult.out.trim())
        if (!isNaN(hwVRAM) && hwVRAM > 512) info.vramMB = hwVRAM
      }
    }

    // ── Method 4: REG_BINARY manual byte parse ────────────────────────────────
    // HardwareInformation.MemorySize is stored as REG_BINARY (raw bytes, little-endian)
    // Standard parseInt() fails on binary data — must read bytes and reconstruct uint64
    if (info.vramMB < 512) {
      const regBinResult = await runPS(`
        $base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
        $found = $false
        Get-ChildItem $base -EA SilentlyContinue | Where-Object { $_.PSChildName -match '^\\d{4}$' } | ForEach-Object {
          if ($found) { return }
          try {
            $raw = (Get-ItemProperty -Path $_.PSPath -EA SilentlyContinue).'HardwareInformation.MemorySize'
            if ($raw -is [byte[]]) {
              $bytes = $raw
              while ($bytes.Length -lt 8) { $bytes += [byte]0 }
              $val = [System.BitConverter]::ToUInt64($bytes, 0)
              if ($val -gt 536870912) { Write-Output $val; $found = $true }
            } elseif ($raw -is [long] -or $raw -is [int] -or $raw -is [uint32] -or $raw -is [uint64]) {
              if ([long]$raw -gt 536870912) { Write-Output ([long]$raw); $found = $true }
            }
          } catch {}
        }
      `, 15000)
      const regBytes = parseFloat(regBinResult.out.trim())
      if (!isNaN(regBytes) && regBytes > 536870912) info.vramMB = Math.floor(regBytes / (1024 ** 2))
    }

    emitSpecProgress(info.vramMB >= 512 ? 'VRAM detected' : 'VRAM detection complete',
      info.vramMB >= 512 ? `${info.vramMB >= 1024 ? (info.vramMB/1024).toFixed(0)+'GB' : info.vramMB+'MB'} VRAM` : 'VRAM not detected')
    await new Promise(r => setTimeout(r, 50))

    // ── Store HWiNFO availability flag for renderer ───────────────────────────
    const hwInfoCheck = await runPS(
      'if (Get-Process -Name "HWiNFO64","HWiNFO32" -ErrorAction SilentlyContinue) { Write-Output "running" } elseif (Get-Command "HWiNFO64.exe" -ErrorAction SilentlyContinue) { Write-Output "installed" } else { Write-Output "none" }'
    )
    info.hwInfoStatus = hwInfoCheck.out.trim() // 'running' | 'installed' | 'none'

    // ── NVMe detection — robust multi-method ─────────────────────────────────
    emitSpecProgress('Detecting storage…', '')
    // Method 1: Get-PhysicalDisk with both BusType integer AND string matching
    const diskR = await runPS(`
      Get-PhysicalDisk | ForEach-Object {
        $bt = $_.BusType
        $mt = $_.MediaType
        $fn = $_.FriendlyName
        Write-Output "BusType=$bt|MediaType=$mt|FriendlyName=$fn"
      }
    `)
    for (const line of diskR.out.split('\n')) {
      const bt = line.match(/BusType=(\w+)/)?.[1] || ''
      const mt = line.match(/MediaType=(\w+)/)?.[1] || ''
      const fn = line.match(/FriendlyName=(.+)/)?.[1] || ''
      // BusType 17 = NVMe, or string "NVMe", or friendly name contains NVMe keywords
      if (bt === '17' || /nvme/i.test(bt) || /nvme/i.test(mt) || /nvme|m\.2|pcie ssd/i.test(fn)) {
        info.nvme = true; info.diskType = 'NVMe SSD'; break
      }
      if (/ssd/i.test(mt) || /ssd/i.test(fn)) info.diskType = 'SSD'
    }

    // Method 2: Check Win32_DiskDrive model names for NVMe keywords
    if (!info.nvme) {
      const driveR = await runPS(
        'Get-CimInstance Win32_DiskDrive | Select-Object -ExpandProperty Model'
      )
      for (const line of driveR.out.split('\n')) {
        if (/nvme|m\.2|samsung 9\d\d|wd_black sn|ct\d+p\d+|sabrent|firecuda 5/i.test(line)) {
          info.nvme = true; info.diskType = 'NVMe SSD'; break
        }
      }
    }

    // Method 3: Check for StorNVMe driver (definitive — if this driver is loaded, NVMe exists)
    if (!info.nvme) {
      const nvmeDrvR = await runPS(
        'Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.DeviceName -like "*NVM*" -or $_.InfName -like "*nvme*"} | Select-Object -First 1 -ExpandProperty DeviceName'
      )
      if (nvmeDrvR.ok && /nvm/i.test(nvmeDrvR.out)) {
        info.nvme = true; info.diskType = 'NVMe SSD'
      }
    }

    // Method 4: Registry StorNVMe service — if it exists, NVMe is present
    if (!info.nvme) {
      const nvmeRegR = await runPS(
        'Test-Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\stornvme"'
      )
      if (nvmeRegR.out.trim().toLowerCase() === 'true') {
        info.nvme = true; info.diskType = 'NVMe SSD'
      }
    }

    // Disk usage
    const diskUsageR = await runPS('$d=Get-PSDrive C; Write-Output ($d.Used); Write-Output ($d.Used+$d.Free)')
    const du = diskUsageR.out.split('\n').map(l => parseInt(l.trim())).filter(n => !isNaN(n))
    if (du.length >= 2) {
      info.diskUsedGB = Math.floor(du[0] / (1024**3))
      info.diskTotalGB = Math.floor(du[1] / (1024**3))
    }

    emitSpecProgress('Storage detected', `${info.diskType} · ${info.diskUsedGB}/${info.diskTotalGB} GB`)
    await new Promise(r => setTimeout(r, 50))

    // OS caption + extra specs for dashboard
    emitSpecProgress('Detecting OS info…', '')
    const extraR = await runPS(`
$cpu   = Get-CimInstance Win32_Processor | Select-Object -First 1
$os    = Get-CimInstance Win32_OperatingSystem
$nic   = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true } | Select-Object -First 1
$nicPh = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.Name -eq $nic.Description } | Select-Object -First 1
$gpu   = Get-CimInstance Win32_VideoController | Select-Object -First 1
$ram   = Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1
$sys   = Get-CimInstance Win32_ComputerSystem
Write-Output "CPU_THREADS=$($cpu.NumberOfLogicalProcessors)"
Write-Output "OS_CAPTION=$($os.Caption)"
Write-Output "NIC_NAME=$($nic.Description)"
Write-Output "NIC_SPEED=$($nicPh.Speed)"
Write-Output "GPU_DRIVER=$($gpu.DriverVersion)"
Write-Output "RAM_SPEED=$($ram.Speed)"
Write-Output "RAM_TYPE=$($ram.MemoryType)"
Write-Output "COMP_NAME=$($sys.DNSHostName)"
Write-Output "DOMAIN=$($sys.PartOfDomain)"
`)
    for (const line of extraR.out.split('\n')) {
      const [key, ...rest] = line.trim().split('=')
      const val = rest.join('=').trim()
      if (key === 'CPU_THREADS' && parseInt(val)) info.cpuThreads = parseInt(val)
      if (key === 'OS_CAPTION' && val) info.osCaption = val
      if (key === 'NIC_NAME' && val) info.nicName = val
      if (key === 'NIC_SPEED' && parseInt(val)) info.nicSpeed = Math.round(parseInt(val) / 1e9)
      if (key === 'GPU_DRIVER' && val) info.gpuDriver = val
      if (key === 'RAM_SPEED' && parseInt(val)) info.ramHz = parseInt(val)
      if (key === 'RAM_TYPE') {
        const types = {20:'DDR',21:'DDR2',22:'DDR2 FB',24:'DDR3',26:'DDR4',34:'DDR5'}
        info.ramType = types[parseInt(val)] || ''
      }
      if (key === 'COMP_NAME' && val) info.computerName = val
    }
    // Build ramSpec string: "2×DDR5-5600 · 32 GB" style
    if (info.ramType && info.ramHz) {
      const sticks = Math.max(1, Math.round((info.ramGB || 0) / 16)) // rough guess
      info.ramSpec = `${info.ramType}-${info.ramHz}`
    }

    // ── Wi-Fi detection ───────────────────────────────────────────────────────
    // Check if the primary active network adapter is wireless (802.11)
    emitSpecProgress('Detecting network type…', '')
    const wifiR = await runPS(`
      $wifi = Get-NetAdapter | Where-Object {
        $_.Status -eq 'Up' -and (
          $_.PhysicalMediaType -eq 'Native 802.11' -or
          $_.PhysicalMediaType -eq 'Wireless LAN' -or
          $_.InterfaceDescription -like '*Wireless*' -or
          $_.InterfaceDescription -like '*Wi-Fi*' -or
          $_.InterfaceDescription -like '*802.11*' -or
          $_.Name -like '*Wi-Fi*' -or
          $_.Name -like '*Wireless*'
        )
      } | Select-Object -First 1
      if ($wifi) { Write-Output "WIFI_YES=$($wifi.InterfaceDescription)" } else { Write-Output "WIFI_NO" }
    `, 10000)
    const wifiLine = wifiR.out.trim()
    if (wifiLine.startsWith('WIFI_YES=')) {
      info.isWifi = true
      info.wifiAdapter = wifiLine.replace('WIFI_YES=', '').trim()
    }

    // ── Laptop detection ──────────────────────────────────────────────────────
    // Check for battery — most reliable laptop indicator
    const laptopR = await runPS(
      'if (Get-CimInstance Win32_Battery -EA SilentlyContinue) { Write-Output "LAPTOP" } else { Write-Output "DESKTOP" }',
      8000
    )
    info.isLaptop = laptopR.out.trim() === 'LAPTOP'
    emitSpecProgress('Network detected', info.isWifi ? `Wi-Fi (${info.wifiAdapter || 'wireless'})` : 'Ethernet')
    await new Promise(r => setTimeout(r, 50))

    emitSpecProgress('Checking admin rights…', '')
    info.isAdmin = await checkAdmin()
    emitSpecProgress('Done', info.isAdmin ? 'Running as Administrator' : 'Not running as Administrator')
    await new Promise(r => setTimeout(r, 50))
  }

  info.fivemInstalled = fs.existsSync(path.join(process.env.LOCALAPPDATA || '', 'FiveM', 'FiveM.app'))

  return info
}

// ─── IPC Handlers ─────────────────────────────────────────────────────────────

ipcMain.handle('get-system-info', async () => {
  const info = await detectSystemInfo()
  // Cache fields for webhook use (webhookLaunch fires before detectSystemInfo completes)
  if (info.gpu)      getSystemFields._cachedGpu      = info.gpu
  if (info.isLaptop) getSystemFields._cachedIsLaptop = info.isLaptop
  if (info.isWifi)   getSystemFields._cachedIsWifi   = info.isWifi
  return info
})

ipcMain.handle('is-admin', async () => checkAdmin())

ipcMain.handle('relaunch-admin', () => {
  if (IS_WIN) {
    const { execSync } = require('child_process')
    try {
      execSync(`powershell -Command "Start-Process '${process.execPath}' -Verb RunAs"`, { windowsHide: true })
      app.quit()
    } catch {}
  }
})

ipcMain.handle('window-minimize', (e) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (win) { win.minimize() }  // triggers the 'minimize' event which hides to tray
})
ipcMain.handle('window-maximize', (e) => {
  const win = BrowserWindow.fromWebContents(e.sender)
  if (win?.isMaximized()) win.restore()
  else win?.maximize()
})
ipcMain.handle('window-close', (e) => BrowserWindow.fromWebContents(e.sender)?.close())

ipcMain.handle('load-settings', () => loadSettings())
ipcMain.handle('save-settings', (_, data) => { saveSettings(data); return true })

ipcMain.handle('open-external', (_, url) => shell.openExternal(url))

// ─── Discord RPC update handlers ──────────────────────────────────────────────
ipcMain.handle('discord-set-page', (_, page) => {
  discordCurrentPage = page
  if (!discordActivityOverride) updateDiscordPresence()
})
ipcMain.handle('discord-set-tweaks', (_, fps) => {
  discordFpsGained = fps
  if (!discordActivityOverride) updateDiscordPresence()
})
ipcMain.handle('discord-set-applied-count', (_, count) => {
  discordTweaksApplied = count
  if (!discordActivityOverride) updateDiscordPresence()
})
ipcMain.handle('discord-set-fivem', (_, hasFiveM) => {
  discordHasFiveM = hasFiveM
  updateDiscordPresence()
})
ipcMain.handle('discord-set-pulse-game', (_, gameName) => {
  discordPulseGame = gameName  // null to clear
  updateDiscordPresence()
})
ipcMain.handle('discord-set-active-game', (_, gameName) => {
  discordActiveGame = gameName  // null to clear
  updateDiscordPresence()
})
ipcMain.handle('discord-op-start', (_, { type, step }) => {
  const phrases = type === 'health' ? HEALTH_PHRASES
    : type === 'preflight' ? SCAN_PHRASES
    : type === 'autoop' ? OPTI_PHRASES
    : SCAN_PHRASES
  setDiscordOverride(phrases, step || '')
})
ipcMain.handle('discord-op-step', (_, step) => {
  updateDiscordOverrideState(step)
})
ipcMain.handle('discord-op-end', () => {
  clearDiscordOverride()
})

// ─── Tweak execution handler ──────────────────────────────────────────────────
ipcMain.handle('run-tweak', async (_, { id, action }) => {
  const send = (msg, level = 'info') => {
    mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString(), source: id })
  }

  send(`Running: ${id} [${action}]`, 'head')

  try {
    const result = await TWEAKS[id]?.[action]?.(send, runPS, runCmd, regAdd, regDelete)
    return { ok: true, result }
  } catch (e) {
    send(`Error in ${id}: ${e.message}`, 'err')
    return { ok: false, error: e.message }
  }
})

// ─── Real-time metrics ────────────────────────────────────────────────────────
let _netPrevBytes = null
let _cpuTempFailed = false  // skip expensive WMI calls after confirmed unavailable
ipcMain.handle('get-metrics', async () => {
  const metrics = {
    cpu: 0, ram: 0, ramUsed: 0, ramTotal: 0,
    disk: 0, diskUsed: 0, diskTotal: 0,
    gpu: 0, netDown: 0, netUp: 0,
    uptime: 0, processes: 0,
    cpuTemp: 0, gpuTemp: 0
  }

  const cpuTempScript = _cpuTempFailed ? `Write-Output "CPU_TEMP=-1"` : `
$cpu_temp = -1
# LibreHardwareMonitor WMI (most accurate, if running)
try {
  $sensors = Get-WmiObject -Namespace "root\\LibreHardwareMonitor" -Class "Sensor" -ErrorAction Stop |
    Where-Object { $_.SensorType -eq "Temperature" -and $_.Name -match "CPU Package|Core #0|Core Average" }
  if ($sensors) {
    $best = ($sensors | Sort-Object Value -Descending | Select-Object -First 1).Value
    if ($best -gt 20 -and $best -lt 120) { $cpu_temp = [math]::Round($best, 1) }
  }
} catch {}
if ($cpu_temp -eq -1) {
  try {
    $sensors = Get-WmiObject -Namespace "root\\OpenHardwareMonitor" -Class "Sensor" -ErrorAction Stop |
      Where-Object { $_.SensorType -eq "Temperature" -and $_.Name -match "CPU Package|Core #0" }
    if ($sensors) {
      $best = ($sensors | Sort-Object Value -Descending | Select-Object -First 1).Value
      if ($best -gt 20 -and $best -lt 120) { $cpu_temp = [math]::Round($best, 1) }
    }
  } catch {}
}
if ($cpu_temp -eq -1) {
  try {
    $zones = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop
    $hotZone = $zones | Sort-Object CurrentTemperature -Descending | Select-Object -First 1
    if ($hotZone) {
      $t = [math]::Round($hotZone.CurrentTemperature / 10 - 273.15, 1)
      if ($t -gt 20 -and $t -lt 120) { $cpu_temp = $t }
    }
  } catch {}
}
Write-Output "CPU_TEMP=$cpu_temp"`

  // Use tagged key=value output to avoid any line-count fragility
  const script = `
$cpu = (Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
$os  = Get-CimInstance Win32_OperatingSystem
$ram_total = $os.TotalVisibleMemorySize * 1024
$ram_free  = $os.FreePhysicalMemory * 1024
$ram_used  = $ram_total - $ram_free
$d = Get-PSDrive C
$disk_used  = $d.Used
$disk_total = $d.Used + $d.Free
$uptime_sec = [int](New-TimeSpan -Start $os.LastBootUpTime -End (Get-Date)).TotalSeconds
$procs = (Get-Process).Count
$net = Get-NetAdapterStatistics -ErrorAction SilentlyContinue | Where-Object { $_.ReceivedBytes -gt 0 } | Select-Object -First 1
$net_rb = if ($net) { $net.ReceivedBytes } else { 0 }
$net_sb = if ($net) { $net.SentBytes } else { 0 }
Write-Output "CPU=$cpu"
Write-Output "RAM_TOTAL=$ram_total"
Write-Output "RAM_USED=$ram_used"
Write-Output "DISK_USED=$disk_used"
Write-Output "DISK_TOTAL=$disk_total"
Write-Output "UPTIME=$uptime_sec"
Write-Output "PROCS=$procs"
Write-Output "NET_RB=$net_rb"
Write-Output "NET_SB=$net_sb"
`
  const r = await runPS(script + '\n' + cpuTempScript)
  const kv = {}
  for (const line of r.out.split('\n')) {
    const eq = line.indexOf('=')
    if (eq > 0) kv[line.slice(0, eq).trim()] = line.slice(eq + 1).trim()
  }
  const kn = (k) => parseFloat(kv[k]) || 0

  metrics.cpu       = Math.round(kn('CPU'))
  const ramTotal    = kn('RAM_TOTAL'), ramUsed = kn('RAM_USED')
  metrics.ramTotal  = Math.round(ramTotal / (1024**3))
  metrics.ramUsed   = Math.round(ramUsed  / (1024**3))
  metrics.ram       = ramTotal > 0 ? Math.round((ramUsed / ramTotal) * 100) : 0
  const diskUsed    = kn('DISK_USED'), diskTotal = kn('DISK_TOTAL')
  metrics.diskUsed  = Math.round(diskUsed  / (1024**3))
  metrics.diskTotal = Math.round(diskTotal / (1024**3))
  metrics.disk      = diskTotal > 0 ? Math.round((diskUsed / diskTotal) * 100) : 0
  metrics.uptime    = Math.round(kn('UPTIME'))
  metrics.processes = Math.round(kn('PROCS'))
  const rawCpuTemp = parseFloat(kv['CPU_TEMP'])
  if (rawCpuTemp === -1) _cpuTempFailed = true
  metrics.cpuTemp = (rawCpuTemp > 0) ? rawCpuTemp : 0
  const netRb = kn('NET_RB'), netSb = kn('NET_SB')
  if (_netPrevBytes && netRb > 0) {
    metrics.netDown = Math.max(0, Math.round((netRb - _netPrevBytes.rb) / 3))
    metrics.netUp   = Math.max(0, Math.round((netSb - _netPrevBytes.sb) / 3))
  }
  _netPrevBytes = { rb: netRb, sb: netSb }

  // GPU load — separate call using PDH counters (Get-Counter outputs multi-line, must be isolated)
  try {
    const gpuR = await runPS(`
$g = 0
try {
  $g = [math]::Round((Get-Counter "\GPU Engine(*engtype_3D)\Utilization Percentage" -EA Stop).CounterSamples |
    Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum)
} catch {
  try {
    $g = [math]::Round((Get-Counter "\GPU Engine(*)\Utilization Percentage" -EA Stop).CounterSamples |
      Where-Object { $_.InstanceName -notlike '*videodecode*' -and $_.InstanceName -notlike '*videoprocessing*' } |
      Measure-Object -Property CookedValue -Sum | Select-Object -ExpandProperty Sum)
  } catch { $g = 0 }
}
Write-Output "GPU_LOAD=$([int][math]::Min(100,[math]::Max(0,$g)))"
`)
    const gMatch = gpuR.out.match(/GPU_LOAD=(\d+)/)
    if (gMatch) metrics.gpu = parseInt(gMatch[1])
  } catch {}

  // GPU temp + usage via nvidia-smi (returns "temp, gpu%" on one line)
  try {
    const nvR = await runCmd('nvidia-smi --query-gpu=temperature.gpu,utilization.gpu --format=csv,noheader,nounits')
    if (nvR.ok && nvR.out.trim()) {
      const parts = nvR.out.trim().split(',').map(s => parseFloat(s.trim()))
      if (!isNaN(parts[0]) && parts[0] > 0) metrics.gpuTemp = parts[0]
      if (!isNaN(parts[1]) && metrics.gpu === 0) metrics.gpu = Math.round(parts[1])
    }
  } catch {}

  // Fallback GPU temp — OpenHardwareMonitor WMI (works for AMD too)
  if (metrics.gpuTemp === 0) {
    try {
      const ohmR = await runPS(`try { Get-WmiObject -Namespace "root/OpenHardwareMonitor" -Class Sensor -EA Stop | Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -like '*GPU*' } | Select-Object -First 1 -ExpandProperty Value } catch { Write-Output 0 }`)
      const t = parseFloat(ohmR.out.trim())
      if (t > 0) metrics.gpuTemp = t
    } catch {}
  }

  return metrics
})

// ─── Restore Points ───────────────────────────────────────────────────────────
ipcMain.handle('create-restore-point', async (_, desc) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  send('Creating System Restore Point…', 'head')

  // Enable SR + override frequency
  await runPS("Enable-ComputerRestore -Drive 'C:\\' -ErrorAction SilentlyContinue")
  await regAdd('HKLM', 'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SystemRestore', 'SystemRestorePointCreationFrequency', 'REG_DWORD', '0')

  const r = await runPS(`Checkpoint-Computer -Description "${desc || 'JylliTool_Restore'}" -RestorePointType MODIFY_SETTINGS -ErrorAction Stop`)
  if (r.ok) { send('✓ Restore Point created!', 'ok'); return { ok: true } }

  // WMI fallback
  const r2 = await runPS(`(Get-WmiObject -Class SystemRestore -Namespace root\\default).CreateRestorePoint("${desc || 'JylliTool_Restore'}",12,100)`)
  if (r2.ok) { send('✓ Restore Point created (WMI)!', 'ok'); return { ok: true } }

  send(`Restore Point failed: ${r.err}`, 'err')
  send('Tip: Run as Admin and ensure C:\\ drive protection is ON.', 'info')
  return { ok: false, error: r.err }
})

ipcMain.handle('list-restore-points', async () => {
  const r = await runPS('Get-ComputerRestorePoint | Sort-Object SequenceNumber -Descending | Select-Object SequenceNumber,Description,CreationTime | ConvertTo-Json')
  if (!r.ok || !r.out) return []
  try {
    const data = JSON.parse(r.out)
    return Array.isArray(data) ? data : [data]
  } catch { return [] }
})

ipcMain.handle('restore-to-point', async (_, seq) => {
  const r = await runPS(`Restore-Computer -RestorePoint ${seq} -Confirm:$false -ErrorAction Stop`)
  if (!r.ok) { shell.openExternal('rstrui.exe'); return { ok: false } }
  return { ok: true }
})

// ─── Debloat: scan installed apps ─────────────────────────────────────────────
ipcMain.handle('scan-uwp-apps', async () => {
  const r = await runPS('Get-AppxPackage -AllUsers | Select-Object -ExpandProperty Name | ConvertTo-Json')
  if (!r.ok || !r.out) return []
  try { return JSON.parse(r.out) } catch { return r.out.split('\n').map(l => l.trim()).filter(Boolean) }
})

ipcMain.handle('remove-uwp-apps', async (_, packages) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  const results = []
  send(`Removing ${packages.length} UWP app(s)…`, 'head')
  for (let i = 0; i < packages.length; i++) {
    const pkg = packages[i]
    mainWindow?.webContents.send('debloat-progress', { current: i + 1, total: packages.length, pkg })
    const r = await runPS(`Get-AppxPackage -Name '${pkg}' -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue`)
    send(`  [${i+1}/${packages.length}] ${pkg}: ${r.ok ? 'removed' : 'not found/skipped'}`, r.ok ? 'ok' : 'info')
    results.push({ pkg, ok: r.ok })
  }
  send('UWP removal complete.', 'ok')
  return results
})

ipcMain.handle('disable-services', async (_, services) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  for (const svc of services) {
    await runCmd(`sc config ${svc} start= disabled`)
    await runCmd(`sc stop ${svc}`)
    send(`  ${svc}: disabled`, 'ok')
  }
  return { ok: true }
})

ipcMain.handle('disable-scheduled-tasks', async (_, tasks) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  for (const task of tasks) {
    const r = await runPS(`Disable-ScheduledTask -TaskPath '${task.path}' -TaskName '${task.name}' -ErrorAction SilentlyContinue`)
    send(`  ${task.name}: ${r.ok ? 'disabled' : 'skipped/not found'}`, r.ok ? 'ok' : 'warn')
  }
  return { ok: true }
})

ipcMain.handle('get-unused-devices', async () => {
  // Use pnputil to enumerate all devices including hidden/ghost ones
  const r = await runPS(`
    $env:DEVMGR_SHOW_NONPRESENT_DEVICES = '1'
    Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
      ($_.Status -eq 'Error' -or $_.Status -eq 'Unknown' -or $_.Present -eq $false) -and
      $_.FriendlyName -ne $null -and $_.FriendlyName -ne '' -and
      $_.Class -notin @('Volume','DiskDrive','CDRom','Processor','Computer','System')
    } | ForEach-Object {
      Write-Output "NAME=$($_.FriendlyName)|ID=$($_.InstanceId)|STATUS=$($_.Status)|CLASS=$($_.Class)|PRESENT=$($_.Present)"
    }
  `, 20000)
  const devices = []
  for (const line of r.out.split('\n')) {
    const name = line.match(/NAME=([^|]+)/)?.[1]?.trim()
    const id = line.match(/ID=([^|]+)/)?.[1]?.trim()
    const status = line.match(/STATUS=([^|]+)/)?.[1]?.trim()
    const cls = line.match(/CLASS=([^|]+)/)?.[1]?.trim()
    const present = line.match(/PRESENT=([^|]+)/)?.[1]?.trim()
    if (name && id) devices.push({ name, id, status, cls, present: present === 'True' })
  }
  return devices
})

ipcMain.handle('remove-devices', async (_, ids) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  for (const id of ids) {
    // Use pnputil to permanently remove the device node (works on ghost/non-present devices)
    const r = await runPS(`
      $result = & pnputil /remove-device "${id}" 2>&1
      if ($LASTEXITCODE -eq 0 -or $result -match 'successfully') {
        Write-Output "OK"
      } else {
        # Fallback: Remove-PnpDevice for present devices
        try {
          Get-PnpDevice | Where-Object {$_.InstanceId -eq '${id}'} | Remove-PnpDevice -Confirm:$false -EA Stop
          Write-Output "OK"
        } catch { Write-Output "FAIL" }
      }
    `)
    const ok = r.out.includes('OK')
    send(`  ${ok ? 'Removed' : 'Could not remove'}: ${id.split('\\').pop()}`, ok ? 'ok' : 'warn')
  }
  return { ok: true }
})

// ─── FiveM in-game settings (gta5_settings.xml) ──────────────────────────────
// Format: <KeyName value="val" /> inside <graphics> block
const FIVEM_SETTINGS_DEFS = [
  { key: 'LodScale',            label: 'LOD Scale',              impact: 'high',      rec: '1.000000',  type: 'select', opts: [['0.500000','0.5 (Best FPS)'],['1.000000','1.0 (Default)'],['1.500000','1.5'],['2.000000','2.0 (High)']] },
  { key: 'PedLodBias',          label: 'Ped LOD Bias',           impact: 'high',      rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Best FPS)'],['0.200000','Low'],['0.500000','Medium'],['1.000000','High']] },
  { key: 'VehicleLodBias',      label: 'Vehicle LOD Bias',       impact: 'high',      rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Best FPS)'],['0.200000','Low'],['0.500000','Medium'],['1.000000','High']] },
  { key: 'ShadowQuality',       label: 'Shadow Quality',         impact: 'very high', rec: '1',         type: 'select', opts: [['0','Off'],['1','Normal (Recommended)'],['2','High'],['3','Very High'],['4','Ultra']] },
  { key: 'ReflectionQuality',   label: 'Reflection Quality',     impact: 'medium',    rec: '1',         type: 'select', opts: [['0','Off'],['1','Normal (Recommended)'],['2','High'],['3','Very High'],['4','Ultra']] },
  { key: 'ReflectionMSAA',      label: 'Reflection MSAA',        impact: 'medium',    rec: '0',         type: 'select', opts: [['0','Off (Recommended)'],['2','2x'],['4','4x'],['8','8x']] },
  { key: 'MSAA',                label: 'MSAA Anti-Aliasing',     impact: 'high',      rec: '0',         type: 'select', opts: [['0','Off (Recommended)'],['2','2x'],['4','4x'],['8','8x']] },
  { key: 'FXAA_Enabled',        label: 'FXAA Anti-Aliasing',     impact: 'low',       rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'TXAA_Enabled',        label: 'TXAA Anti-Aliasing',     impact: 'medium',    rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'ParticleQuality',     label: 'Particle Quality',       impact: 'medium',    rec: '1',         type: 'select', opts: [['0','Low'],['1','Normal (Recommended)'],['2','High'],['3','Very High'],['4','Ultra']] },
  { key: 'WaterQuality',        label: 'Water Quality',          impact: 'medium',    rec: '0',         type: 'select', opts: [['0','Low (Recommended)'],['1','Normal'],['2','High'],['3','Very High'],['4','Ultra']] },
  { key: 'GrassQuality',        label: 'Grass Quality',          impact: 'very high', rec: '0',         type: 'select', opts: [['0','Off (Best FPS)'],['1','Normal'],['2','High'],['3','Ultra']] },
  { key: 'ShaderQuality',       label: 'Shader Quality',         impact: 'high',      rec: '2',         type: 'select', opts: [['0','Low'],['1','Normal'],['2','High (Recommended)'],['3','Very High'],['4','Ultra']] },
  { key: 'Shadow_Distance',     label: 'Shadow Draw Distance',   impact: 'high',      rec: '0.500000',  type: 'select', opts: [['0.250000','0.25 (Best FPS)'],['0.500000','0.5 (Recommended)'],['0.750000','0.75'],['1.000000','1.0 (Max)']] },
  { key: 'Shadow_SoftShadows',  label: 'Soft Shadows',           impact: 'medium',    rec: '0',         type: 'select', opts: [['0','Off (Recommended)'],['1','Softer'],['2','Softest'],['3','AMD CHS']] },
  { key: 'UltraShadows_Enabled',label: 'Ultra Shadows',          impact: 'high',      rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'Shadow_ParticleShadows',label:'Particle Shadows',      impact: 'medium',    rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'Shadow_LongShadows',  label: 'Long Shadows',           impact: 'low',       rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'CityDensity',         label: 'Ped Density',            impact: 'very high', rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Best FPS)'],['0.500000','Low (Recommended)'],['1.000000','Normal'],['1.500000','High']] },
  { key: 'VehicleVarietyMultiplier', label: 'Vehicle Density',   impact: 'very high', rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Best FPS)'],['0.500000','Low (Recommended)'],['1.000000','Normal'],['1.500000','High']] },
  { key: 'PedVarietyMultiplier',label: 'Ped Variety',            impact: 'high',      rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Best FPS)'],['0.500000','Low'],['1.000000','Normal']] },
  { key: 'PostFX',              label: 'Post FX Quality',        impact: 'medium',    rec: '1',         type: 'select', opts: [['0','Off'],['1','Normal (Recommended)'],['2','High'],['3','Very High'],['4','Ultra']] },
  { key: 'DoF',                 label: 'Depth of Field',         impact: 'medium',    rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'MotionBlurStrength',  label: 'Motion Blur',            impact: 'low',       rec: '0.000000',  type: 'select', opts: [['0.000000','Off (Recommended)'],['0.500000','Low'],['1.000000','Full']] },
  { key: 'Tessellation',        label: 'Tessellation',           impact: 'medium',    rec: '0',         type: 'select', opts: [['0','Off (Recommended)'],['1','Normal'],['2','High'],['3','Very High']] },
  { key: 'AnisotropicFiltering',label: 'Anisotropic Filtering',  impact: 'low',       rec: '16',        type: 'select', opts: [['0','Off'],['2','2x'],['4','4x'],['8','8x'],['16','16x (Recommended)']] },
  { key: 'SSAO',                label: 'Ambient Occlusion (SSAO)',impact: 'medium',   rec: '0',         type: 'select', opts: [['0','Off (Recommended)'],['1','Normal'],['2','High']] },
  { key: 'Lighting_FogVolumes', label: 'Volumetric Fog',         impact: 'medium',    rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
  { key: 'HdStreamingInFlight', label: 'HD Streaming While Flying', impact: 'high',   rec: 'false',     type: 'select', opts: [['false','Off (Recommended)'],['true','On']] },
]

// gta5_settings.xml finder — primary path is %APPDATA%\CitizenFX\gta5_settings.xml
function findFivemSettings(manualPath) {
  if (manualPath && manualPath.trim()) {
    try { require('fs').accessSync(manualPath.trim()); return manualPath.trim() } catch {}
    return null
  }
  const primary = path.join(process.env.APPDATA || '', 'CitizenFX', 'gta5_settings.xml')
  try { require('fs').accessSync(primary); return primary } catch {}
  return null
}

ipcMain.handle('fivem-read-settings', async (_, manualPath) => {
  const settingsPath = findFivemSettings(manualPath)
  if (!settingsPath) {
    return {
      ok: false,
      reason: 'gta5_settings.xml not found at %APPDATA%\\CitizenFX\\gta5_settings.xml',
      hint: 'Launch FiveM, open Settings → Graphics, adjust any setting and close FiveM. The file will be created automatically.'
    }
  }
  let raw
  try { raw = require('fs').readFileSync(settingsPath, 'utf8') } catch (e) {
    return { ok: false, reason: `Could not read ${settingsPath}: ${e.message}` }
  }
  const vals = {}
  for (const def of FIVEM_SETTINGS_DEFS) {
    // Format: <KeyName value="val" />  or  <KeyName value="val">
    const m = raw.match(new RegExp(`<${def.key}\\s+value="([^"]*)"`, 'i'))
    vals[def.key] = m ? m[1] : null
  }
  return { ok: true, vals, defs: FIVEM_SETTINGS_DEFS, path: settingsPath }
})

ipcMain.handle('fivem-write-settings', async (_, changes, manualPath) => {
  const settingsPath = findFivemSettings(manualPath)
  if (!settingsPath) return { ok: false, reason: 'gta5_settings.xml not found — cannot save.' }
  let content
  try { content = require('fs').readFileSync(settingsPath, 'utf8') } catch (e) {
    return { ok: false, reason: `Could not read file: ${e.message}` }
  }

  // Backup before write
  try { require('fs').writeFileSync(settingsPath + '.bak', content, 'utf8') } catch (_) {}

  for (const [key, val] of Object.entries(changes)) {
    // Exact-match regex — word boundary after key name prevents partial matches
    // Matches: <KeyName value="old" />  or  <KeyName value="old">
    const existing = new RegExp(`(<${key}(?=\\s)\\s+value=")[^"]*("\\s*/?>)`, 'i')
    if (existing.test(content)) {
      content = content.replace(existing, `$1${val}$2`)
    } else {
      // Key not present — inject it inside the <graphics> block
      content = content.replace(/(<\/graphics>)/i, `    <${key} value="${val}" />\n  $1`)
    }
  }

  try {
    require('fs').writeFileSync(settingsPath, content, 'utf8')
  } catch (e) {
    return { ok: false, reason: `Could not write file: ${e.message}` }
  }

  // Verify: re-read and confirm a sample key landed correctly
  try {
    const verify = require('fs').readFileSync(settingsPath, 'utf8')
    const firstKey = Object.keys(changes)[0]
    const firstVal = changes[firstKey]
    const check = new RegExp(`<${firstKey}\\s+value="${firstVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i')
    if (!check.test(verify)) {
      return { ok: false, reason: 'File written but verification failed — values may not have saved correctly.' }
    }
  } catch (_) {}

  return { ok: true, path: settingsPath, changed: Object.keys(changes).length }
})

// ─── FiveM cache clear ────────────────────────────────────────────────────────
ipcMain.handle('fivem-clear-cache', async () => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  const lad = process.env.LOCALAPPDATA || ''
  const fivemBase = path.join(lad, 'FiveM', 'FiveM.app')
  const paths = [
    path.join(fivemBase, 'data', 'cache'),
    path.join(fivemBase, 'data', 'server-cache'),
    path.join(fivemBase, 'data', 'server-cache-priv'),
    path.join(lad, 'NVIDIA', 'DXCache'),
  ]
  let total = 0
  for (const p of paths) {
    if (fs.existsSync(p)) {
      const r = await runPS(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue; (Get-ChildItem -Path '${p}' -ErrorAction SilentlyContinue).Count`)
      send(`  ${path.basename(p)}: cleared`, 'ok')
      total++
    }
  }
  send(`Cache cleared — ${total} locations processed.`, 'ok')
  return { ok: true }
})

// ─── Auto-Optimization ────────────────────────────────────────────────────────
ipcMain.handle('run-auto-opti', async (_, sysInfo) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  const progress = (step, total, label) => mainWindow?.webContents.send('auto-opti-progress', { step, total, label })

  setDiscordOverride(OPTI_PHRASES, 'Starting…')

  send('◆ Auto-Optimization started', 'head')
  progress(0, 1, 'Creating Restore Point…')
  const rp = await createRestorePointInternal(send)
  if (!rp.ok) {
    clearDiscordOverride()
    send('✗ Aborting — no tweaks were applied. Fix the restore point issue and try again.', 'err')
    progress(0, 1, 'Aborted — restore point failed')
    return { ok: false, aborted: true, reason: rp.reason }
  }

  // If the wizard passed a selectedTweaks list, run only those via TWEAKS registry.
  // Fall back to the legacy buildAutoOptiSteps for backwards compatibility.
  const selected = sysInfo?.selectedTweaks
  if (selected && Array.isArray(selected) && selected.length > 0) {
    for (let i = 0; i < selected.length; i++) {
      const id = selected[i]
      const tweak = TWEAKS[id]
      progress(i + 1, selected.length, id)
      updateDiscordOverrideState(`Step ${i + 1} / ${selected.length}`)
      send(`  [${i+1}/${selected.length}] ${id}`, 'info')
      if (tweak?.apply) {
        try {
          await tweak.apply(
            (msg, lvl) => send(`    ${msg}`, lvl || 'ok'),
            runPS, runCmd, regAdd, regDelete
          )
        } catch (e) { send(`    Skipped: ${e.message}`, 'warn') }
      } else {
        send(`    No handler for ${id} — skipped`, 'warn')
      }
    }
  } else {
    const steps = buildAutoOptiSteps(sysInfo)
    for (let i = 0; i < steps.length; i++) {
      const { label, fn } = steps[i]
      progress(i + 1, steps.length, label)
      send(`  [${i+1}/${steps.length}] ${label}`, 'info')
      try { await fn(send, runPS, runCmd, regAdd, regDelete) } catch (e) { send(`    Skipped: ${e.message}`, 'warn') }
    }
  }

  send('✓ Auto-Optimization complete! Reboot recommended.', 'ok')
  progress(999, 999, 'Complete!')
  clearDiscordOverride()

  // Analytics
  const tweakCount     = selected?.length || 0
  const includedFiveM  = (selected || []).some(id => id.startsWith('fivem-'))
  const fpsEst         = sysInfo?.fpsEstimate || 0
  sessionTweaksCount  += tweakCount
  webhookAutoOpti(tweakCount, fpsEst, includedFiveM)

  return { ok: true }
})

async function createRestorePointInternal(send) {
  // Enable System Restore on C: and remove the 24-hour cooldown throttle
  await runPS("Enable-ComputerRestore -Drive 'C:\\' -ErrorAction SilentlyContinue")
  await regAdd('HKLM', 'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SystemRestore', 'SystemRestorePointCreationFrequency', 'REG_DWORD', '0')

  // Attempt to create the restore point — treat any error as a hard failure
  const r = await runPS(`
    try {
      Checkpoint-Computer -Description "JylliTool_AutoOpti" -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
      Write-Output "RP_OK"
    } catch {
      Write-Output "RP_FAIL:$($_.Exception.Message)"
    }
  `)
  const out = r.out?.trim() || ''
  if (out.startsWith('RP_OK') || r.ok) {
    send('✓ Restore Point created — system is safe to optimise.', 'ok')
    return { ok: true }
  }
  const reason = out.replace('RP_FAIL:', '') || r.err || 'Unknown error'
  send(`✗ Restore Point FAILED: ${reason}`, 'err')
  return { ok: false, reason }
}

function buildAutoOptiSteps(si) {
  const steps = []
  const add = (label, fn) => steps.push({ label, fn })

  add('Disable Game DVR / Xbox recording', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_Enabled -Value 0 -Force')
    await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" -Force | Out-Null; Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" -Name AllowGameDVR -Value 0 -Force')
    s('Game DVR disabled.', 'ok')
  })

  add('Set Visual Effects → Best Performance', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" -Name VisualFXSetting -Value 2 -Force')
    s('Visual effects minimised.', 'ok')
  })

  add('MMCSS GPU Priority=8, Games scheduling=High', async (s, ps) => {
    const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
    // 10, not 0 — value 0 starves USB HID interrupts (causes ghost mouse clicks and audio beeps)
    await ps(`Set-ItemProperty -Path "${base}" -Name SystemResponsiveness -Value 10 -Force`)
    await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "GPU Priority" -Value 8 -Force`)
    await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Priority" -Value 6 -Force`)
    await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Scheduling Category" -Value "High" -Force`)
    s('MMCSS applied.', 'ok')
  })

  add('Enable GPU Hardware Scheduling', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 2 -Force')
    s('GPU HW Scheduling enabled. Reboot required.', 'ok')
  })

  add('Disable CPU core parking', async (s, _, cmd) => {
    await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR CPMINCORES 100')
    await cmd('powercfg /setactive SCHEME_CURRENT')
    s('CPU core parking disabled.', 'ok')
  })

  add('NTFS: disable 8.3 names + last-access timestamps', async (s, _, cmd) => {
    await cmd('fsutil behavior set disable8dot3 1')
    await cmd('fsutil behavior set disablelastaccess 1')
    s('NTFS tweaks applied.', 'ok')
  })

  add('Enable MSI Interrupt Mode for PCI devices', async (s, ps) => {
    await ps(`
      $base = "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\PCI"
      $netGuid  = '{4d36e972-e325-11ce-bfc1-08002be10318}'
      $count = 0; $skipped = 0
      Get-ChildItem $base -EA SilentlyContinue | ForEach-Object {
        Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
          $props   = Get-ItemProperty -Path $_.PSPath -EA SilentlyContinue
          $classGuid = $props.ClassGUID
          $service   = $props.Service
          $isNet   = $classGuid -eq $netGuid
          $isWifi  = $service -match 'iwifi|netathr|bcmwl|rt[l6]8|rtswlan|Netwtw|IntelWifi|netrtwlane|RtlW|mrvlpcie|rtwlane'
          if ($isNet -or $isWifi) { $skipped++; return }
          $msi = "$($_.PSPath)\\Device Parameters\\Interrupt Management\\MessageSignaledInterruptProperties"
          New-Item -Path $msi -Force -EA SilentlyContinue | Out-Null
          Set-ItemProperty -Path $msi -Name MSISupported -Value 1 -Force -EA SilentlyContinue
          $count++
        }
      }
      Write-Output "MSI set on $count devices, skipped $skipped network adapters"
    `)
    s('MSI mode applied to PCI devices (network adapters excluded).', 'ok')
  })

  add('Disable WPAD (Web Proxy Auto-Discovery)', async (s, _, cmd) => {
    await cmd('sc config WinHttpAutoProxySvc start= disabled')
    await cmd('sc stop WinHttpAutoProxySvc')
    s('WPAD disabled.', 'ok')
  })

  add('Disable NetBIOS over TCP/IP', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters" -Name NetbiosOptions -Value 2 -Force')
    s('NetBIOS disabled.', 'ok')
  })

  add('Disable network throttling', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" -Name NetworkThrottlingIndex -Value 0xFFFFFFFF -Force')
    s('Network throttling disabled.', 'ok')
  })

  add('Optimise TCP stack', async (s, _, cmd) => {
    await cmd('netsh int tcp set global autotuninglevel=normal')
    await cmd('netsh int tcp set global rss=enabled')
    await cmd('netsh int tcp set global chimney=disabled')
    await cmd('netsh int tcp set global ecncapability=disabled')
    await cmd('netsh int tcp set global timestamps=disabled')
    s('TCP stack optimised.', 'ok')
  })

  add('Remove QoS 20% bandwidth reservation', async (s, ps) => {
    await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" -Force | Out-Null; Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" -Name NonBestEffortLimit -Value 0 -Force')
    s('QoS reservation removed.', 'ok')
  })

  add('Disable startup app delay', async (s, ps) => {
    await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" -Force | Out-Null; Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" -Name StartupDelayInMSec -Value 0 -Force')
    s('Startup delay removed.', 'ok')
  })

  add('Disable Windows tips & suggestions', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" -Name SoftLandingEnabled -Value 0 -Force -ErrorAction SilentlyContinue')
    await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager" -Name "SubscribedContent-338389Enabled" -Value 0 -Force -ErrorAction SilentlyContinue')
    s('Windows tips disabled.', 'ok')
  })

  add('Set foreground process priority boost', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" -Name Win32PrioritySeparation -Value 38 -Force')
    s('Priority separation set.', 'ok')
  })

  if (!si?.isLaptop) {
    add('Import Ultimate Performance power plan', async (s, _, cmd) => {
      const r = await cmd('powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61')
      const m = r.out.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i)
      if (m) { await cmd(`powercfg -setactive ${m[1]}`); s('Ultimate Performance activated.', 'ok') }
      else s('Could not import power plan (may already exist).', 'warn')
    })
  }

  if (si?.nvme) {
    add('NVMe StorPort latency tweak', async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name TelemetryPerformanceHighResolutionTimer -Value 0 -Force -ErrorAction SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name BusyPauseTime -Value 0 -Force -ErrorAction SilentlyContinue')
      s('NVMe StorPort tweak applied.', 'ok')
    })
  }

  if (si?.ramGB >= 16) {
    add('Disable SysMain (SuperFetch)', async (s, _, cmd) => {
      await cmd('sc stop SysMain')
      await cmd('sc config SysMain start= disabled')
      s('SysMain disabled.', 'ok')
    })
  }

  add('Disable telemetry & DiagTrack service', async (s, ps, cmd) => {
    await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Force -EA SilentlyContinue | Out-Null')
    await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Name AllowTelemetry -Value 0 -Force')
    await cmd('sc config DiagTrack start= disabled')
    await cmd('sc stop DiagTrack')
    await cmd('sc config dmwappushservice start= disabled')
    await cmd('sc stop dmwappushservice')
    s('Telemetry & DiagTrack disabled.', 'ok')
  })

  add('Disable Power Throttling', async (s, ps) => {
    await ps('New-Item -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling" -Force -EA SilentlyContinue | Out-Null')
    await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling" -Name PowerThrottlingOff -Value 1 -Force')
    s('Power Throttling disabled.', 'ok')
  })

  add('Disable UWP background apps', async (s, ps) => {
    await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" -Force -EA SilentlyContinue | Out-Null')
    await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" -Name GlobalUserDisabled -Value 1 -Force')
    s('UWP background apps disabled.', 'ok')
  })

  add('Disable mouse acceleration', async (s, ps) => {
    await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseSpeed -Value "0" -Force')
    await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold1 -Value "0" -Force')
    await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold2 -Value "0" -Force')
    s('Mouse acceleration disabled.', 'ok')
  })

  add('Disable Windows Update P2P delivery optimization', async (s, ps) => {
    await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Force -EA SilentlyContinue | Out-Null')
    await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Name DODownloadMode -Value 0 -Force')
    s('Delivery Optimization disabled.', 'ok')
  })

  add('Disable automatic maintenance', async (s, ps, cmd) => {
    await ps('New-Item -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" -Force -EA SilentlyContinue | Out-Null')
    await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" -Name MaintenanceDisabled -Value 1 -Force')
    await cmd('schtasks /Change /TN "\\Microsoft\\Windows\\TaskScheduler\\Regular Maintenance" /Disable 2>nul')
    s('Automatic maintenance disabled.', 'ok')
  })

  if (!si?.isLaptop) {
    add('Disable USB Selective Suspend', async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('USB Selective Suspend disabled.', 'ok')
    })

    add('Disable PCIe Link State Power Management', async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 0')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('PCIe power management disabled.', 'ok')
    })
  }

  const ethernetFilter = `$_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and $_.PhysicalMediaType -ne 'Wireless LAN' -and $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and $_.InterfaceDescription -notlike '*802.11*' -and $_.InterfaceDescription -notlike '*Virtual*' -and $_.InterfaceDescription -notlike '*Loopback*'`

  add('Disable NIC interrupt moderation (Ethernet only)', async (s, ps) => {
    await ps(`
      Get-NetAdapter | Where-Object {${ethernetFilter}} | ForEach-Object {
        $n = $_.Name
        foreach ($prop in @("Interrupt Moderation","InterruptModeration","Interrupt Moderation Rate")) {
          try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; break } catch {}
        }
      }
    `)
    s('NIC interrupt moderation disabled.', 'ok')
  })

  add('Disable NIC flow control (Ethernet only)', async (s, ps) => {
    await ps(`
      Get-NetAdapter | Where-Object {${ethernetFilter}} | ForEach-Object {
        $n = $_.Name
        foreach ($prop in @("Flow Control","FlowControl")) {
          try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; break } catch {}
        }
      }
    `)
    s('NIC flow control disabled.', 'ok')
  })

  add('Disable Energy Efficient Ethernet (Ethernet only)', async (s, ps) => {
    await ps(`
      Get-NetAdapter | Where-Object {${ethernetFilter}} | ForEach-Object {
        $n = $_.Name
        foreach ($prop in @("Energy Efficient Ethernet","EEE","Green Ethernet","Energy-Efficient Ethernet")) {
          try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; break } catch {}
        }
      }
    `)
    s('Energy Efficient Ethernet disabled.', 'ok')
  })

  return steps
}

// ─── Individual tweak registry ────────────────────────────────────────────────
// Each tweak has apply/restore functions
const TWEAKS = {
  // ── Windows Optimizations ────────────────────────────────────────────────
  'game-dvr': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_Enabled -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_FSEBehaviorMode -Value 2 -Force')
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" -Force -ErrorAction SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" -Name AllowGameDVR -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\SOFTWARE\\Microsoft\\GameBar" -Name UseNexusForGameBarEnabled -Value 0 -Force -ErrorAction SilentlyContinue')
      s('Game DVR / Xbox Game Bar disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_Enabled -Value 1 -Force')
      s('Game DVR restored.', 'ok')
    }
  },
  'telemetry': {
    apply: async (s, ps, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Force -ErrorAction SilentlyContinue | Out-Null; Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Name AllowTelemetry -Value 0 -Force')
      await cmd('sc config DiagTrack start= disabled'); await cmd('sc stop DiagTrack')
      await cmd('sc config dmwappushservice start= disabled'); await cmd('sc stop dmwappushservice')
      s('Telemetry & DiagTrack disabled.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config DiagTrack start= auto'); await cmd('sc start DiagTrack')
      s('Telemetry restored.', 'ok')
    }
  },
  'visual-effects': {
    apply: async (s, ps, cmd) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" -Name VisualFXSetting -Value 2 -Force')
      await cmd('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v TaskbarAnimations /t REG_DWORD /d 0 /f')
      s('Visual effects → Best Performance.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" -Name VisualFXSetting -Value 0 -Force')
      s('Visual effects restored.', 'ok')
    }
  },
  'mmcss': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
      // SystemResponsiveness=10 (not 0!) — 0 starves USB HID driver causing mouse ghosting and audio beeps
      await ps(`Set-ItemProperty -Path "${base}" -Name SystemResponsiveness -Value 10 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "GPU Priority" -Value 8 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Priority" -Value 6 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Scheduling Category" -Value "High" -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "SFIO Priority" -Value "High" -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Clock Rate" -Value 10000 -Force`)
      s('MMCSS game scheduling applied. GPU Priority=8, games=High.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
      await ps(`Set-ItemProperty -Path "${base}" -Name SystemResponsiveness -Value 20 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "GPU Priority" -Value 2 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Priority" -Value 2 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Scheduling Category" -Value "Medium" -Force`)
      s('MMCSS defaults restored.', 'ok')
    }
  },
  'sysmain': {
    apply: async (s, _, cmd) => { await cmd('sc stop SysMain'); await cmd('sc config SysMain start= disabled'); s('SysMain disabled.', 'ok') },
    restore: async (s, _, cmd) => { await cmd('sc config SysMain start= auto'); await cmd('sc start SysMain'); s('SysMain restored.', 'ok') }
  },
  'power-throttling': {
    apply: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling" -Name PowerThrottlingOff -Value 1 -Force'); s('Power throttling disabled.', 'ok') },
    restore: async (s, ps) => { await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling" -Name PowerThrottlingOff -ErrorAction SilentlyContinue'); s('Power throttling restored.', 'ok') }
  },
  'hpet': {
    apply: async (s, _, cmd) => {
      await cmd('bcdedit /deletevalue useplatformclock')
      await cmd('bcdedit /set disabledynamictick yes')
      await cmd('bcdedit /set useplatformtick yes')
      s('HPET disabled. Reboot required.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('bcdedit /set useplatformclock true')
      await cmd('bcdedit /deletevalue disabledynamictick')
      s('HPET restored. Reboot required.', 'ok')
    }
  },
  'tsc-sync': {
    apply: async (s, _, cmd) => { await cmd('bcdedit /set tscsyncpolicy enhanced'); s('TSC sync → Enhanced. Reboot required.', 'ok') },
    restore: async (s, _, cmd) => { await cmd('bcdedit /set tscsyncpolicy default'); s('TSC sync restored.', 'ok') }
  },
  'gpu-hwsch': {
    apply: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 2 -Force'); s('GPU HW Scheduling enabled. Reboot required.', 'ok') },
    restore: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 1 -Force'); s('GPU HW Scheduling disabled.', 'ok') }
  },
  'ntfs': {
    apply: async (s, _, cmd) => {
      await cmd('fsutil behavior set disable8dot3 1')
      await cmd('fsutil behavior set disablelastaccess 1')
      await cmd('fsutil behavior set encryptpagingfile 0')
      s('NTFS performance tweaks applied.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('fsutil behavior set disable8dot3 0')
      await cmd('fsutil behavior set disablelastaccess 0')
      s('NTFS defaults restored.', 'ok')
    }
  },
  'msi-mode': {
    apply: async (s, ps) => {
      await ps(`
        $base = "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\PCI"
        $count = 0; $skipped = 0
        Get-ChildItem $base -EA SilentlyContinue | ForEach-Object {
          Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
            $devPath = $_.PSPath
            # Skip network/Wi-Fi adapters — Intel/Qualcomm/Realtek Wi-Fi drivers
            # do not support MSI correctly and will drop the adapter after reboot
            $classGuid = (Get-ItemProperty -Path $devPath -EA SilentlyContinue).ClassGUID
            $service   = (Get-ItemProperty -Path $devPath -EA SilentlyContinue).Service
            $isNet = $classGuid -eq '{4d36e972-e325-11ce-bfc1-08002be10318}'
            $isWifi = $service -match 'iwifi|netathr|bcmwl|rt[l6]8|rtswlan|netadap|Netwtw|netrtwlane|RtlW|IntelWifi|Netwtw0[2-9]|mrvlpcie|rtwlane'
            if ($isNet -or $isWifi) { $skipped++; return }
            $p = "$devPath\\Device Parameters\\Interrupt Management\\MessageSignaledInterruptProperties"
            New-Item -Path $p -Force -EA SilentlyContinue | Out-Null
            Set-ItemProperty -Path $p -Name MSISupported -Value 1 -Force -EA SilentlyContinue
            $count++
          }
        }
        Write-Output "Set on $count devices, skipped $skipped network adapters"
      `)
      s('MSI Interrupt Mode applied to PCI devices (Wi-Fi/NICs excluded). Reboot required.', 'ok')
    },
    restore: async (s) => s('Use Device Manager to revert MSI mode per device if needed.', 'info')
  },
  'memory-compression': {
    apply: async (s, ps) => { const r = await ps('Disable-MMAgent -MemoryCompression'); s(r.ok ? 'Memory compression disabled.' : `Error: ${r.err}`, r.ok ? 'ok' : 'err') },
    restore: async (s, ps) => { await ps('Enable-MMAgent -MemoryCompression'); s('Memory compression re-enabled.', 'ok') }
  },
  'bcd-tweaks': {
    apply: async (s, _, cmd) => {
      await cmd('bcdedit /debug off')
      await cmd('bcdedit /timeout 0')
      await cmd('bcdedit /set {default} bootlog no')
      s('BCD boot tweaks applied. Reboot required.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('bcdedit /set {default} recoveryenabled yes')
      await cmd('bcdedit /timeout 30')
      s('BCD tweaks restored.', 'ok')
    }
  },
  'spectre-meltdown': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name FeatureSettingsOverride -Value 3 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name FeatureSettingsOverrideMask -Value 3 -Force')
      s('CPU mitigations disabled. SECURITY RISK. Reboot required.', 'warn')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name FeatureSettingsOverride -EA SilentlyContinue')
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name FeatureSettingsOverrideMask -EA SilentlyContinue')
      s('CPU mitigations restored. Reboot required.', 'ok')
    }
  },
  'cortana': {
    apply: async (s, ps, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCortana -Value 0 -Force')
      await cmd('sc config WSearch start= disabled'); await cmd('sc stop WSearch')
      s('Cortana + Windows Search indexer disabled.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config WSearch start= auto'); await cmd('sc start WSearch')
      s('Cortana/WSearch restored.', 'ok')
    }
  },
  'hibernate': {
    apply: async (s, _, cmd) => { const r = await cmd('powercfg /hibernate off'); s(r.ok ? 'Hibernate disabled.' : `Error: ${r.err}`, r.ok ? 'ok' : 'err') },
    restore: async (s, _, cmd) => { await cmd('powercfg /hibernate on'); s('Hibernate re-enabled.', 'ok') }
  },
  'explorer-perf': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" -Name LaunchTo -Value 1 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" -Name StartupDelayInMSec -Value 0 -Force -ErrorAction SilentlyContinue')
      s('Explorer performance tweaks applied.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" -Name StartupDelayInMSec -EA SilentlyContinue')
      s('Explorer tweaks removed.', 'ok')
    }
  },
  'win-tips': {
    apply: async (s, ps) => {
      const base = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager'
      for (const name of ['SoftLandingEnabled','SubscribedContent-338389Enabled','SubscribedContent-310093Enabled']) {
        await ps(`Set-ItemProperty -Path "${base}" -Name "${name}" -Value 0 -Force -EA SilentlyContinue`)
      }
      s('Windows tips & suggestions disabled.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager'
      for (const name of ['SoftLandingEnabled','SubscribedContent-338389Enabled']) {
        await ps(`Set-ItemProperty -Path "${base}" -Name "${name}" -Value 1 -Force -EA SilentlyContinue`)
      }
      s('Windows tips restored.', 'ok')
    }
  },
  'priority-sep': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" -Name Win32PrioritySeparation -Value 38 -Force')
      // 10 not 0 — 0 starves USB HID interrupts
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" -Name SystemResponsiveness -Value 10 -Force')
      s('Foreground process priority boost applied.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" -Name Win32PrioritySeparation -Value 2 -Force')
      s('Priority separation restored.', 'ok')
    }
  },
  'tdr-delay': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrLevel -Value 3 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrDelay -Value 10 -Force')
      s('GPU TDR delay extended to 10s. Reduces driver timeout crashes.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrLevel -EA SilentlyContinue')
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrDelay -EA SilentlyContinue')
      s('TDR settings restored.', 'ok')
    }
  },
  'disable-fso': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_DXGIHonorFSEWindowsCompatible -Value 1 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_FSEBehaviorMode -Value 2 -Force')
      s('Fullscreen optimizations disabled globally.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_DXGIHonorFSEWindowsCompatible -EA SilentlyContinue')
      s('Fullscreen optimizations restored.', 'ok')
    }
  },

  'disable-background-apps': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" -Name GlobalUserDisabled -Value 1 -Force')
      s('UWP background apps disabled globally.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" -Name GlobalUserDisabled -Value 0 -Force -EA SilentlyContinue')
      s('UWP background apps restored.', 'ok')
    }
  },
  'disable-auto-maintenance': {
    apply: async (s, ps, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" -Name MaintenanceDisabled -Value 1 -Force')
      await cmd('schtasks /Change /TN "\\Microsoft\\Windows\\TaskScheduler\\Regular Maintenance" /Disable 2>nul')
      s('Automatic maintenance disabled.', 'ok')
    },
    restore: async (s, ps, cmd) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance" -Name MaintenanceDisabled -Force -EA SilentlyContinue')
      await cmd('schtasks /Change /TN "\\Microsoft\\Windows\\TaskScheduler\\Regular Maintenance" /Enable 2>nul')
      s('Automatic maintenance re-enabled.', 'ok')
    }
  },
  'disable-mouse-accel': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseSpeed -Value "0" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold1 -Value "0" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold2 -Value "0" -Force')
      s('Mouse acceleration (Enhance Pointer Precision) disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseSpeed -Value "1" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold1 -Value "6" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold2 -Value "10" -Force')
      s('Mouse acceleration restored to Windows default.', 'ok')
    }
  },
  'disable-delivery-opt': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Name DODownloadMode -Value 0 -Force')
      s('Windows Update P2P delivery optimization disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Name DODownloadMode -Force -EA SilentlyContinue')
      s('Delivery Optimization restored to default.', 'ok')
    }
  },
  'usb-suspend': {
    apply: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('USB Selective Suspend disabled. No more USB stutter or disconnects.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 1')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 1')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('USB Selective Suspend restored.', 'ok')
    }
  },

  // ── Hardware ──────────────────────────────────────────────────────────────
  'nvidia-max-perf': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerEnable -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerLevel -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerLevelAC -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PerfLevelSrc -Value 8738 -Force -EA SilentlyContinue`)
      }
      s('NVIDIA Max Performance mode applied.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const name of ['PowerMizerEnable','PowerMizerLevel','PowerMizerLevelAC','PerfLevelSrc']) {
        for (const idx of ['0000','0001']) await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name ${name} -EA SilentlyContinue`)
      }
      s('NVIDIA defaults restored.', 'ok')
    }
  },
  'nvidia-low-latency': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0000'
      await ps(`Set-ItemProperty -Path "${base}" -Name D3PCLatency -Value 1 -Force -EA SilentlyContinue`)
      await ps(`Set-ItemProperty -Path "${base}" -Name F1MaxLatency -Value 0 -Force -EA SilentlyContinue`)
      s('NVIDIA Ultra Low-Latency applied.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0000'
      await ps(`Remove-ItemProperty -Path "${base}" -Name D3PCLatency -EA SilentlyContinue`)
      await ps(`Remove-ItemProperty -Path "${base}" -Name F1MaxLatency -EA SilentlyContinue`)
      s('NVIDIA latency restored.', 'ok')
    }
  },
  'disable-hdcp': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0000" -Name RMHdcpKeyglobZero -Value 1 -Force -EA SilentlyContinue')
      s('HDCP disabled. Reboot required.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0000" -Name RMHdcpKeyglobZero -EA SilentlyContinue')
      s('HDCP restored.', 'ok')
    }
  },
  'ultimate-perf': {
    apply: async (s, _, cmd) => {
      const r = await cmd('powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61')
      const m = r.out.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i)
      if (m) { await cmd(`powercfg -setactive ${m[1]}`); s(`Ultimate Performance activated (${m[1]}).`, 'ok') }
      else s('Could not import plan — it may already exist.', 'warn')
    },
    restore: async (s, _, cmd) => {
      await cmd('powercfg -setactive 381b4222-f694-41f0-9685-ff5bb260df2e')
      s('Balanced power plan activated.', 'ok')
    }
  },
  'cpu-parking': {
    apply: async (s, _, cmd) => { await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR CPMINCORES 100'); await cmd('powercfg /setactive SCHEME_CURRENT'); s('CPU core parking disabled.', 'ok') },
    restore: async (s, _, cmd) => { await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR CPMINCORES 0'); await cmd('powercfg /setactive SCHEME_CURRENT'); s('CPU core parking restored.', 'ok') }
  },
  'cstates': {
    apply: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 100')
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100')
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR IDLEPROMOTE 0')
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR IDLEDEMOTE 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('C-States minimised.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMIN 5')
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PROCTHROTTLEMAX 100')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('C-State defaults restored.', 'ok')
    }
  },
  'nvme-latency': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name TelemetryPerformanceHighResolutionTimer -Value 0 -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name BusyPauseTime -Value 0 -Force -EA SilentlyContinue')
      s('NVMe StorPort latency tweak applied.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name TelemetryPerformanceHighResolutionTimer -EA SilentlyContinue')
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort" -Name BusyPauseTime -EA SilentlyContinue')
      s('NVMe tweaks removed.', 'ok')
    }
  },
  'speed-shift': {
    apply: async (s, _, cmd) => { await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PERFAUTONOMOUS 1'); await cmd('powercfg /setactive SCHEME_CURRENT'); s('Intel Speed Shift enabled (AMD: no effect).', 'ok') },
    restore: async (s, _, cmd) => { await cmd('powercfg /setacvalueindex SCHEME_CURRENT SUB_PROCESSOR PERFAUTONOMOUS 0'); await cmd('powercfg /setactive SCHEME_CURRENT'); s('Speed Shift reverted.', 'ok') }
  },
  'pcie-power': {
    apply: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 0')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('PCIe Link State Power Management disabled. GPU and NVMe stay at full speed.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 2')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 2')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      s('PCIe Link State Power Management restored.', 'ok')
    }
  },

  // ── Network ───────────────────────────────────────────────────────────────
  'tcp-stack': {
    apply: async (s, _, cmd) => {
      await cmd('netsh int tcp set global autotuninglevel=normal')
      await cmd('netsh int tcp set global rss=enabled')
      await cmd('netsh int tcp set global chimney=disabled')
      await cmd('netsh int tcp set global ecncapability=disabled')
      await cmd('netsh int tcp set global timestamps=disabled')
      await cmd('netsh int tcp set global maxsynretransmissions=2')
      await cmd('netsh int tcp set global initialrto=2000')
      s('TCP stack optimised.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('netsh int tcp set global autotuninglevel=normal')
      await cmd('netsh int tcp set global rss=enabled')
      await cmd('netsh int tcp set global chimney=enabled')
      s('TCP defaults restored.', 'ok')
    }
  },
  'qos-reserve': {
    apply: async (s, ps) => { await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" -Name NonBestEffortLimit -Value 0 -Force'); s('QoS 20% reservation removed.', 'ok') },
    restore: async (s, ps) => { await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" -Name NonBestEffortLimit -EA SilentlyContinue'); s('QoS reservation restored.', 'ok') }
  },
  'winsock-reset': {
    apply: async (s, _, cmd) => { await cmd('netsh winsock reset catalog'); await cmd('netsh int ip reset resetlog.txt'); s('Winsock & IP stack reset. REBOOT required.', 'ok') },
    restore: async (s) => s('Winsock reset is permanent — reboot was applied.', 'info')
  },
  'flush-dns': {
    apply: async (s, _, cmd) => { const r = await cmd('ipconfig /flushdns'); s(r.out || 'DNS cache flushed.', 'ok') },
    restore: async (s) => s('DNS flush is one-time only.', 'info')
  },
  'dns-cloudflare': {
    apply: async (s, ps) => {
      const adapters = await ps("Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty Name")
      for (const a of adapters.out.split('\n').map(l => l.trim()).filter(Boolean)) {
        await require('child_process').execFileSync ? null : null
        const cmd2 = require('child_process').execFile
        const r1 = await new Promise(res => require('child_process').exec(`netsh interface ip set dns name="${a}" static 1.1.1.1 primary`, { windowsHide: true }, (e, o) => res({ ok: !e })))
        const r2 = await new Promise(res => require('child_process').exec(`netsh interface ip add dns name="${a}" 1.0.0.1 index=2`, { windowsHide: true }, (e, o) => res({ ok: !e })))
        s(`  ${a}: Cloudflare 1.1.1.1 set`, 'ok')
      }
    },
    restore: async (s, ps) => {
      const adapters = await ps("Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty Name")
      for (const a of adapters.out.split('\n').map(l => l.trim()).filter(Boolean)) {
        await new Promise(res => require('child_process').exec(`netsh interface ip set dns name="${a}" dhcp`, { windowsHide: true }, res))
        s(`  ${a}: DNS restored to DHCP`, 'ok')
      }
    }
  },
  'dns-google': {
    apply: async (s, ps) => {
      const adapters = await ps("Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty Name")
      for (const a of adapters.out.split('\n').map(l => l.trim()).filter(Boolean)) {
        await new Promise(res => require('child_process').exec(`netsh interface ip set dns name="${a}" static 8.8.8.8 primary`, { windowsHide: true }, res))
        await new Promise(res => require('child_process').exec(`netsh interface ip add dns name="${a}" 8.8.4.4 index=2`, { windowsHide: true }, res))
        s(`  ${a}: Google DNS 8.8.8.8 set`, 'ok')
      }
    },
    restore: async (s, ps) => {
      const adapters = await ps("Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty Name")
      for (const a of adapters.out.split('\n').map(l => l.trim()).filter(Boolean)) {
        await new Promise(res => require('child_process').exec(`netsh interface ip set dns name="${a}" dhcp`, { windowsHide: true }, res))
      }
      s('DNS restored to DHCP.', 'ok')
    }
  },
  'disable-netbios': {
    apply: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters" -Name NetbiosOptions -Value 2 -Force'); s('NetBIOS over TCP/IP disabled.', 'ok') },
    restore: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters" -Name NetbiosOptions -Value 0 -Force'); s('NetBIOS restored.', 'ok') }
  },
  'disable-ipv6': {
    apply: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip6\\Parameters" -Name DisabledComponents -Value 255 -Force'); s('IPv6 disabled. Reboot required.', 'ok') },
    restore: async (s, ps) => { await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip6\\Parameters" -Name DisabledComponents -EA SilentlyContinue'); s('IPv6 restored. Reboot required.', 'ok') }
  },
  'disable-wpad': {
    apply: async (s, _, cmd) => { await cmd('sc config WinHttpAutoProxySvc start= disabled'); await cmd('sc stop WinHttpAutoProxySvc'); s('WPAD disabled.', 'ok') },
    restore: async (s, _, cmd) => { await cmd('sc config WinHttpAutoProxySvc start= manual'); await cmd('sc start WinHttpAutoProxySvc'); s('WPAD restored.', 'ok') }
  },
  'net-throttling': {
    apply: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" -Name NetworkThrottlingIndex -Value 0xFFFFFFFF -Force'); s('Network throttling disabled.', 'ok') },
    restore: async (s, ps) => { await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" -Name NetworkThrottlingIndex -Value 10 -Force'); s('Network throttling restored.', 'ok') }
  },
  'tcp-buffers': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters'
      await ps(`Set-ItemProperty -Path "${base}" -Name GlobalMaxTcpWindowSize -Value 65535 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TcpWindowSize -Value 65535 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name Tcp1323Opts -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name SackOpts -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name DefaultTTL -Value 64 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TcpAckFrequency -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TCPNoDelay -Value 1 -Force`)
      s('TCP buffers + Nagle disabled applied.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters'
      for (const n of ['GlobalMaxTcpWindowSize','TcpWindowSize','TcpAckFrequency','TCPNoDelay']) await ps(`Remove-ItemProperty -Path "${base}" -Name ${n} -EA SilentlyContinue`)
      s('TCP buffer tweaks removed.', 'ok')
    }
  },
  'nic-offloads': {
    apply: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and
          $_.PhysicalMediaType -ne 'Native 802.11' -and
          $_.PhysicalMediaType -ne 'Wireless LAN' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and
          $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*802.11*' -and
          $_.InterfaceDescription -notlike '*Virtual*' -and
          $_.InterfaceDescription -notlike '*Loopback*'
        } | ForEach-Object {
          Disable-NetAdapterChecksumOffload -Name $_.Name -ErrorAction SilentlyContinue
          Disable-NetAdapterLso -Name $_.Name -ErrorAction SilentlyContinue
          Disable-NetAdapterRsc -Name $_.Name -ErrorAction SilentlyContinue
          Write-Output "Disabled offloads: $($_.Name)"
        }
      `)
      s('NIC hardware offloads disabled (Ethernet only — Wi-Fi skipped).', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and
          $_.PhysicalMediaType -ne 'Native 802.11' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and
          $_.InterfaceDescription -notlike '*Wi-Fi*'
        } | ForEach-Object {
          Enable-NetAdapterChecksumOffload -Name $_.Name -EA SilentlyContinue
          Enable-NetAdapterLso -Name $_.Name -EA SilentlyContinue
        }
      `)
      s('NIC offloads restored.', 'ok')
    }
  },

  'nic-interrupt-mod': {
    apply: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and
          $_.PhysicalMediaType -ne 'Native 802.11' -and $_.PhysicalMediaType -ne 'Wireless LAN' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*802.11*' -and $_.InterfaceDescription -notlike '*Virtual*' -and
          $_.InterfaceDescription -notlike '*Loopback*'
        } | ForEach-Object {
          $n = $_.Name
          $done = $false
          foreach ($prop in @("Interrupt Moderation","InterruptModeration","Interrupt Moderation Rate")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; $done = $true; Write-Output "  $n: Interrupt Moderation disabled"; break } catch {}
          }
          if (-not $done) { Write-Output "  $n: property not found (NIC may not support it)" }
        }
      `)
      s('NIC interrupt moderation disabled (Ethernet only). More consistent ping under load.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*Virtual*'
        } | ForEach-Object {
          $n = $_.Name
          foreach ($prop in @("Interrupt Moderation","InterruptModeration","Interrupt Moderation Rate")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Enabled" -EA Stop; break } catch {}
          }
        }
      `)
      s('NIC interrupt moderation restored.', 'ok')
    }
  },
  'nic-flow-control': {
    apply: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and
          $_.PhysicalMediaType -ne 'Native 802.11' -and $_.PhysicalMediaType -ne 'Wireless LAN' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*802.11*' -and $_.InterfaceDescription -notlike '*Virtual*' -and
          $_.InterfaceDescription -notlike '*Loopback*'
        } | ForEach-Object {
          $n = $_.Name
          foreach ($prop in @("Flow Control","FlowControl")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; Write-Output "  $n: Flow Control disabled"; break } catch {}
          }
        }
      `)
      s('NIC flow control disabled (Ethernet only). Reduces jitter under burst traffic.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*Virtual*'
        } | ForEach-Object {
          $n = $_.Name
          foreach ($prop in @("Flow Control","FlowControl")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Rx & Tx Enabled" -EA Stop; break } catch {}
          }
        }
      `)
      s('NIC flow control restored.', 'ok')
    }
  },
  'nic-energy-efficient': {
    apply: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and
          $_.PhysicalMediaType -ne 'Native 802.11' -and $_.PhysicalMediaType -ne 'Wireless LAN' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*802.11*' -and $_.InterfaceDescription -notlike '*Virtual*' -and
          $_.InterfaceDescription -notlike '*Loopback*'
        } | ForEach-Object {
          $n = $_.Name
          foreach ($prop in @("Energy Efficient Ethernet","EEE","Green Ethernet","Energy-Efficient Ethernet","Ultra Low Power Mode")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Disabled" -EA Stop; Write-Output "  $n: EEE disabled"; break } catch {}
          }
        }
      `)
      s('Energy Efficient Ethernet disabled (Ethernet only). NIC stays at full speed.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`
        Get-NetAdapter | Where-Object {
          $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
          $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*' -and
          $_.InterfaceDescription -notlike '*Virtual*'
        } | ForEach-Object {
          $n = $_.Name
          foreach ($prop in @("Energy Efficient Ethernet","EEE","Green Ethernet","Energy-Efficient Ethernet","Ultra Low Power Mode")) {
            try { Set-NetAdapterAdvancedProperty -Name $n -DisplayName $prop -DisplayValue "Enabled" -EA Stop; break } catch {}
          }
        }
      `)
      s('Energy Efficient Ethernet restored.', 'ok')
    }
  },

  // ── Cleanup ───────────────────────────────────────────────────────────────
  'clean-temp': {
    apply: async (s, ps) => {
      const paths = [
        process.env.TEMP, 'C:\\Windows\\Temp', 'C:\\Windows\\Prefetch',
        'C:\\Windows\\SoftwareDistribution\\Download'
      ]
      for (const p of paths.filter(Boolean)) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue`)
        s(`  ${path.basename(p)}: cleared`, 'ok')
      }
      s('System temp cleanup complete.', 'ok')
    },
    restore: async (s) => s('Temp files cannot be restored — they are expendable.', 'info')
  },
  'clean-browsers': {
    apply: async (s, ps) => {
      const lad = process.env.LOCALAPPDATA || ''
      const caches = {
        'Edge Cache': path.join(lad, 'Microsoft','Edge','User Data','Default','Cache'),
        'Chrome Cache': path.join(lad, 'Google','Chrome','User Data','Default','Cache'),
        'Firefox Cache': path.join(lad, 'Mozilla','Firefox','Profiles'),
      }
      for (const [name, p] of Object.entries(caches)) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue`)
        s(`  ${name}: cleared`, 'ok')
      }
      s('Browser caches cleared.', 'ok')
    },
    restore: async (s) => s('Browser caches cannot be restored.', 'info')
  },
  'clean-discord': {
    apply: async (s, ps) => {
      const apd = process.env.APPDATA || ''
      for (const p of ['Cache','Code Cache','GPUCache'].map(n => path.join(apd,'discord',n))) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue`)
      }
      s('Discord cache cleared.', 'ok')
    },
    restore: async (s) => s('Discord cache cannot be restored.', 'info')
  },
  'clean-nvidia': {
    apply: async (s, ps) => {
      const lad = process.env.LOCALAPPDATA || ''
      for (const p of ['DXCache','GLCache','ComputeCache'].map(n => path.join(lad,'NVIDIA',n))) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue`)
        s(`  NVIDIA ${path.basename(p)}: cleared`, 'ok')
      }
      await ps(`Remove-Item -Path '${path.join(lad,'D3DSCache')}\\*' -Recurse -Force -EA SilentlyContinue`)
      s('NVIDIA shader caches cleared.', 'ok')
    },
    restore: async (s) => s('Shader caches cannot be restored — they rebuild on next launch.', 'info')
  },

  'clean-winsxs': {
    apply: async (s, ps) => {
      s('Running DISM component store cleanup — this can take several minutes…', 'info')
      await ps('Dism.exe /Online /Cleanup-Image /StartComponentCleanup 2>$null', 300000)
      s('WinSxS cleanup complete.', 'ok')
    },
    restore: async (s) => s('Component store cleanup cannot be reversed.', 'info')
  },

  'clean-eventlogs': {
    apply: async (s, ps) => {
      await ps('Get-EventLog -LogName * -EA SilentlyContinue | ForEach-Object { Clear-EventLog -LogName $_.Log -EA SilentlyContinue }')
      s('All Windows event logs cleared.', 'ok')
    },
    restore: async (s) => s('Event logs cannot be restored — entries are gone permanently.', 'info')
  },

  'clean-errorlogs': {
    apply: async (s, ps) => {
      const paths = [
        'C:\\Windows\\Minidump',
        path.join(process.env.LOCALAPPDATA || '', 'CrashDumps'),
        path.join(process.env.APPDATA || '', 'Microsoft', 'Windows', 'WER', 'ReportQueue'),
        path.join(process.env.APPDATA || '', 'Microsoft', 'Windows', 'WER', 'ReportArchive'),
      ]
      for (const p of paths) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -EA SilentlyContinue`)
      }
      s('Crash dumps and error reports cleared.', 'ok')
    },
    restore: async (s) => s('Crash dumps cannot be restored.', 'info')
  },

  'clean-thumbnails': {
    apply: async (s, ps) => {
      // Kill Explorer so thumbcache files can be deleted
      await ps(`
        Stop-Process -Name explorer -Force -EA SilentlyContinue
        Start-Sleep -Milliseconds 800
        Remove-Item -Path "$env:LOCALAPPDATA\\Microsoft\\Windows\\Explorer\\thumbcache_*.db" -Force -EA SilentlyContinue
        Remove-Item -Path "$env:LOCALAPPDATA\\Microsoft\\Windows\\Explorer\\iconcache_*.db" -Force -EA SilentlyContinue
        Start-Process explorer
      `)
      s('Thumbnail and icon cache cleared. Explorer restarted.', 'ok')
    },
    restore: async (s) => s('Caches rebuild automatically — nothing to restore.', 'info')
  },

  'clean-fontcache': {
    apply: async (s, ps) => {
      await ps(`
        Stop-Service -Name FontCache -Force -EA SilentlyContinue
        Remove-Item -Path "$env:LOCALAPPDATA\\Microsoft\\Windows\\FontCache*" -Recurse -Force -EA SilentlyContinue
        Remove-Item -Path "C:\\Windows\\System32\\FNTCACHE.DAT" -Force -EA SilentlyContinue
        Start-Service -Name FontCache -EA SilentlyContinue
      `)
      s('Font cache cleared. Will rebuild on next boot.', 'ok')
    },
    restore: async (s) => s('Font cache rebuilds automatically on next boot.', 'info')
  },

  'clean-browser-history': {
    apply: async (s, ps) => {
      const lad = process.env.LOCALAPPDATA || ''
      const histories = {
        'Edge History': path.join(lad, 'Microsoft', 'Edge', 'User Data', 'Default', 'History'),
        'Edge Cookies': path.join(lad, 'Microsoft', 'Edge', 'User Data', 'Default', 'Cookies'),
        'Chrome History': path.join(lad, 'Google', 'Chrome', 'User Data', 'Default', 'History'),
        'Chrome Cookies': path.join(lad, 'Google', 'Chrome', 'User Data', 'Default', 'Cookies'),
      }
      for (const [name, p] of Object.entries(histories)) {
        await ps(`Remove-Item -Path '${p}' -Force -EA SilentlyContinue`)
        s(`  ${name}: cleared`, 'ok')
      }
      // Firefox profiles — clear places.sqlite
      await ps(`Get-ChildItem "$env:APPDATA\\Mozilla\\Firefox\\Profiles" -Filter "places.sqlite" -Recurse -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue`)
      s('Browser history and cookies cleared.', 'ok')
    },
    restore: async (s) => s('Browser history cannot be restored.', 'info')
  },

  'clean-steam': {
    apply: async (s, ps) => {
      const lad = process.env.LOCALAPPDATA || ''
      const appd = process.env.APPDATA || ''
      const paths = [
        path.join(lad, 'Steam', 'htmlcache'),
        path.join(appd, 'Steam', 'logs'),
        path.join(appd, 'Steam', 'dumps'),
        'C:\\Program Files (x86)\\Steam\\logs',
        'C:\\Program Files (x86)\\Steam\\dumps',
      ]
      for (const p of paths) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -EA SilentlyContinue`)
        s(`  ${path.basename(p)}: cleared`, 'ok')
      }
      s('Steam cache cleared.', 'ok')
    },
    restore: async (s) => s('Steam cache cannot be restored — rebuilds automatically.', 'info')
  },

  'clean-gameservices': {
    apply: async (s, ps) => {
      const lad = process.env.LOCALAPPDATA || ''
      const appd = process.env.APPDATA || ''
      const paths = [
        path.join(lad, 'EA Desktop', 'logs'),
        path.join(lad, 'EA Desktop', 'cache'),
        path.join(appd, 'Battle.net', 'Cache'),
        path.join(appd, 'Battle.net', 'Logs'),
        path.join(lad, 'EpicGamesLauncher', 'Saved', 'Logs'),
        path.join(lad, 'EpicGamesLauncher', 'Saved', 'webcache'),
      ]
      for (const p of paths) {
        await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -EA SilentlyContinue`)
        s(`  ${path.basename(path.dirname(p))} ${path.basename(p)}: cleared`, 'ok')
      }
      s('Game launcher caches cleared.', 'ok')
    },
    restore: async (s) => s('Launcher caches cannot be restored — they rebuild on next launch.', 'info')
  },

  'clean-old-windows': {
    apply: async (s, ps) => {
      if (!fs.existsSync('C:\\Windows.old')) { s('Windows.old not found — nothing to remove.', 'info'); return }
      s('Removing Windows.old — this may take a while…', 'info')
      await ps('Remove-Item -Path "C:\\Windows.old" -Recurse -Force -EA SilentlyContinue', 120000)
      s('Windows.old removed.', 'ok')
    },
    restore: async (s) => s('Windows.old cannot be restored once deleted.', 'info')
  },

  'clean-vs-temp': {
    apply: async (s, ps) => {
      const userProfile = process.env.USERPROFILE || ''
      const paths = [
        path.join(userProfile, '.nuget', 'packages'),
        path.join(userProfile, 'AppData', 'Local', 'Temp', 'NuGetScratch'),
        path.join(process.env.LOCALAPPDATA || '', 'NuGet', 'Cache'),
        path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'TypeScript'),
        path.join(process.env.LOCALAPPDATA || '', 'npm-cache'),
        path.join(process.env.APPDATA || '', 'npm-cache'),
      ]
      for (const p of paths) {
        if (fs.existsSync(p)) {
          await ps(`Remove-Item -Path '${p}\\*' -Recurse -Force -EA SilentlyContinue`)
          s(`  ${path.basename(p)}: cleared`, 'ok')
        }
      }
      s('Dev tool caches cleared.', 'ok')
    },
    restore: async (s) => s('Dev caches cannot be restored — packages re-download on next build.', 'info')
  },

  'fivem-clear-cache': {
    apply: async (s) => {
      const lad = process.env.LOCALAPPDATA || ''
      const fivemBase = path.join(lad, 'FiveM', 'FiveM.app')
      const paths = [
        path.join(fivemBase, 'data', 'cache'),
        path.join(fivemBase, 'data', 'server-cache'),
        path.join(fivemBase, 'data', 'server-cache-priv'),
        path.join(lad, 'NVIDIA', 'DXCache'),
      ]
      for (const p of paths) {
        if (fs.existsSync(p)) {
          await require('fs').promises.rm(p, { recursive: true, force: true }).catch(() => {})
          s(`  ${path.basename(p)}: cleared`, 'ok')
        }
      }
      s('FiveM cache cleared.', 'ok')
    },
    restore: async (s) => s('FiveM cache cannot be restored — rebuilds on next game launch.', 'info')
  },

  // ── FiveM ─────────────────────────────────────────────────────────────────
  'fivem-priority': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
      for (const exe of ['FiveM.exe','GTA5.exe','FiveM_b3095_GTAProcess.exe','CitizenFX.exe']) {
        await ps(`New-Item -Path "${base}\\${exe}\\PerfOptions" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "${base}\\${exe}\\PerfOptions" -Name CpuPriorityClass -Value 3 -Force`)
        s(`  ${exe}: HIGH CPU priority set`, 'ok')
      }
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
      for (const exe of ['FiveM.exe','GTA5.exe','FiveM_b3095_GTAProcess.exe']) {
        await ps(`Set-ItemProperty -Path "${base}\\${exe}\\PerfOptions" -Name CpuPriorityClass -Value 2 -Force -EA SilentlyContinue`)
      }
      s('FiveM/GTA priorities restored to Normal.', 'ok')
    }
  },
  'fivem-io-priority': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
      for (const exe of ['FiveM.exe','GTA5.exe','FiveM_b3095_GTAProcess.exe']) {
        await ps(`New-Item -Path "${base}\\${exe}\\PerfOptions" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "${base}\\${exe}\\PerfOptions" -Name IoPriority -Value 3 -Force`)
        s(`  ${exe}: HIGH I/O priority set`, 'ok')
      }
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
      for (const exe of ['FiveM.exe','GTA5.exe']) await ps(`Set-ItemProperty -Path "${base}\\${exe}\\PerfOptions" -Name IoPriority -Value 2 -Force -EA SilentlyContinue`)
      s('FiveM I/O priority restored.', 'ok')
    }
  },
  'fivem-defender': {
    apply: async (s, ps) => {
      const fivemPath = path.join(process.env.LOCALAPPDATA||'', 'FiveM')
      const r = await ps(`Add-MpPreference -ExclusionPath '${fivemPath}' -ErrorAction SilentlyContinue`)
      s(r.ok ? `Defender exclusion added: ${fivemPath}` : `Could not add exclusion (may need Admin)`, r.ok ? 'ok' : 'warn')
    },
    restore: async (s, ps) => {
      await ps(`Remove-MpPreference -ExclusionPath '${path.join(process.env.LOCALAPPDATA||'','FiveM')}' -EA SilentlyContinue`)
      s('FiveM Defender exclusion removed.', 'ok')
    }
  },
  'fivem-hang-fix': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Desktop" -Name HungAppTimeout -Value 3000 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Desktop" -Name WaitToKillAppTimeout -Value 3000 -Force')
      s('Hang timeout → 3 s. FiveM disconnect hang fixed.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Desktop" -Name HungAppTimeout -Value 5000 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Desktop" -Name WaitToKillAppTimeout -Value 5000 -Force')
      s('Hang timeouts restored to 5 s.', 'ok')
    }
  },
  'fivem-network': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters'
      await ps(`Set-ItemProperty -Path "${base}" -Name TcpAckFrequency -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TCPNoDelay -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TcpDelAckTicks -Value 0 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name DefaultTTL -Value 64 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name MaxUserPort -Value 65534 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name TcpTimedWaitDelay -Value 30 -Force`)
      s('FiveM network tweaks applied. Nagle off, ACK freq=1.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters'
      for (const n of ['TcpAckFrequency','TCPNoDelay','TcpDelAckTicks','TcpTimedWaitDelay']) await ps(`Remove-ItemProperty -Path "${base}" -Name ${n} -EA SilentlyContinue`)
      s('FiveM network tweaks removed.', 'ok')
    }
  },
  'fivem-mmcss': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
      // SystemResponsiveness=10 (not 0!) — 0 starves USB HID causing mouse/input issues
      await ps(`Set-ItemProperty -Path "${base}" -Name SystemResponsiveness -Value 10 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "GPU Priority" -Value 8 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Priority" -Value 6 -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Scheduling Category" -Value "High" -Force`)
      await ps(`Set-ItemProperty -Path "${base}\\Tasks\\Games" -Name "Clock Rate" -Value 10000 -Force`)
      s('FiveM MMCSS scheduling applied. GPU Priority=8, games=High.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" -Name SystemResponsiveness -Value 20 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" -Name "Priority" -Value 2 -Force -EA SilentlyContinue')
      s('MMCSS restored.', 'ok')
    }
  },
  'fivem-fso': {
    apply: async (s, ps) => {
      const compat = 'HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers'
      for (const exe of ['FiveM.exe','GTA5.exe','FiveM_b3095_GTAProcess.exe']) {
        await ps(`New-Item -Path "${compat}" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "${compat}" -Name "${exe}" -Value "~ DISABLEDXMAXIMIZEDWINDOWEDMODE" -Force`)
        s(`  ${exe}: FSO disabled`, 'ok')
      }
    },
    restore: async (s, ps) => {
      const compat = 'HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers'
      for (const exe of ['FiveM.exe','GTA5.exe']) await ps(`Remove-ItemProperty -Path "${compat}" -Name "${exe}" -EA SilentlyContinue`)
      s('Fullscreen optimizations restored.', 'ok')
    }
  },
  'fivem-gpu': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        for (const [n,v] of [['PowerMizerEnable',1],['PowerMizerLevel',1],['PowerMizerLevelAC',1],['PerfLevelSrc',8738],['D3PCLatency',1],['F1MaxLatency',0]]) {
          await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name ${n} -Value ${v} -Force -EA SilentlyContinue`)
        }
      }
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 2 -Force')
      s('NVIDIA max perf + Ultra Low-Latency + HW Scheduling applied.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const n of ['PowerMizerEnable','PowerMizerLevel','D3PCLatency','F1MaxLatency']) for (const idx of ['0000','0001']) await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name ${n} -EA SilentlyContinue`)
      s('GPU settings restored.', 'ok')
    }
  },
  'fivem-vm': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management'
      await ps(`Set-ItemProperty -Path "${base}" -Name DisablePagingExecutive -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name LargeSystemCache -Value 0 -Force`)
      s('Virtual memory optimised for FiveM.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name DisablePagingExecutive -Value 0 -Force')
      s('Virtual memory tweaks restored.', 'ok')
    }
  },
  'fivem-commandline': {
    apply: async (s, ps) => {
      const gtaPathR = await ps(`
        $paths = @(
          "$env:ProgramFiles\\Rockstar Games\\Grand Theft Auto V",
          "$env:ProgramFiles(x86)\\Steam\\steamapps\\common\\Grand Theft Auto V",
          "C:\\Program Files\\Rockstar Games\\Grand Theft Auto V",
          "D:\\Rockstar Games\\Grand Theft Auto V",
          "D:\\SteamLibrary\\steamapps\\common\\Grand Theft Auto V",
          "E:\\SteamLibrary\\steamapps\\common\\Grand Theft Auto V"
        )
        foreach ($p in $paths) { if (Test-Path $p) { Write-Output $p; break } }
      `)
      const gtaPath = gtaPathR.out.trim()
      if (!gtaPath) { s('GTA V install not found — commandline.txt not created.', 'warn'); return }
      const lines = ['-dx11', '-fullscreen', '-notexturebudget', '-high'].join('\r\n')
      await ps(`Set-Content -Path "${gtaPath}\\commandline.txt" -Value "${lines.replace(/\n/g,'`n')}" -Encoding ASCII -Force`)
      s(`GTA V commandline.txt created: DX11, fullscreen, no texture budget, high priority.`, 'ok')
      s(`  Path: ${gtaPath}\\commandline.txt`, 'info')
    },
    restore: async (s, ps) => {
      const gtaPathR = await ps(`
        $paths = @(
          "$env:ProgramFiles\\Rockstar Games\\Grand Theft Auto V",
          "$env:ProgramFiles(x86)\\Steam\\steamapps\\common\\Grand Theft Auto V",
          "C:\\Program Files\\Rockstar Games\\Grand Theft Auto V",
          "D:\\Rockstar Games\\Grand Theft Auto V",
          "D:\\SteamLibrary\\steamapps\\common\\Grand Theft Auto V",
          "E:\\SteamLibrary\\steamapps\\common\\Grand Theft Auto V"
        )
        foreach ($p in $paths) { if (Test-Path $p) { Write-Output $p; break } }
      `)
      const gtaPath = gtaPathR.out.trim()
      if (gtaPath) {
        await ps(`Remove-Item "${gtaPath}\\commandline.txt" -Force -EA SilentlyContinue`)
        s('commandline.txt removed.', 'ok')
      }
    }
  },
  'fivem-streaming-mem': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') {
        s('CitizenFX.ini not found — is FiveM installed?', 'warn'); return
      }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/StreamingMemory\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'StreamingMemory\\s*=\\s*\\d+', 'StreamingMemory=1024' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nStreamingMemory=1024" -Encoding ASCII`)
      }
      s('FiveM StreamingMemory set to 1024 MB — reduces texture pop-in and world stutters.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'StreamingMemory\\s*=\\s*\\d+', 'StreamingMemory=500' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('FiveM StreamingMemory restored to 500 MB.', 'ok')
    }
  },

  // ── FiveM CitizenFX.ini advanced tweaks ──────────────────────────────────
  'fivem-worker-threads': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const threads = require('os').cpus().length
      const val = Math.max(4, Math.min(threads, 16))
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/WorkerThreads\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'WorkerThreads\\s*=\\s*\\d+', 'WorkerThreads=${val}' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nWorkerThreads=${val}" -Encoding ASCII`)
      }
      s(`WorkerThreads set to ${val} (matched to your CPU core count) — faster asset streaming.`, 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'WorkerThreads\\s*=\\s*\\d+', 'WorkerThreads=2' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('WorkerThreads restored to default (2).', 'ok')
    }
  },

  'fivem-disable-crash-reporter': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/DisableCrashReporter\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'DisableCrashReporter\\s*=\\s*\\w+', 'DisableCrashReporter=1' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nDisableCrashReporter=1" -Encoding ASCII`)
      }
      s('Crash reporter disabled — no crash upload overhead on startup or crash.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'DisableCrashReporter\\s*=\\s*\\w+', 'DisableCrashReporter=0' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('Crash reporter re-enabled.', 'ok')
    }
  },

  'fivem-disable-anticheat-upload': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/DisableSteamAchievements\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'DisableSteamAchievements\\s*=\\s*\\w+', 'DisableSteamAchievements=true' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nDisableSteamAchievements=true" -Encoding ASCII`)
      }
      s('Steam Achievements / Rockstar telemetry upload disabled — reduces background traffic.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'DisableSteamAchievements\\s*=\\s*\\w+', 'DisableSteamAchievements=false' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('Steam Achievements restored.', 'ok')
    }
  },

  'fivem-disable-update-checks': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/UpdateChannel\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'UpdateChannel\\s*=\\s*\\w+', 'UpdateChannel=canary' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nUpdateChannel=canary" -Encoding ASCII`)
      }
      s('Update channel set to canary — suppresses automatic update nags on launch.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'UpdateChannel\\s*=\\s*\\w+', 'UpdateChannel=production' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('Update channel restored to production.', 'ok')
    }
  },

  'fivem-preload-ipl': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      if (/MaximumGrass\s*=/i.test(content)) {
        await ps(`(Get-Content "${iniPath}" -Raw) -replace 'MaximumGrass\\s*=\\s*\\d+', 'MaximumGrass=0' | Set-Content "${iniPath}" -Encoding ASCII -Force`)
      } else {
        await ps(`Add-Content -Path "${iniPath}" -Value "\`nMaximumGrass=0" -Encoding ASCII`)
      }
      s('MaximumGrass=0 — removes all grass draw calls, boosts FPS significantly in open areas.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      await ps(`if (Test-Path "${iniPath}") { (Get-Content "${iniPath}" -Raw) -replace 'MaximumGrass\\s*=\\s*\\d+', 'MaximumGrass=60' | Set-Content "${iniPath}" -Encoding ASCII -Force }`)
      s('MaximumGrass restored to default (60).', 'ok')
    }
  },

  'fivem-reduce-draw-distance': {
    apply: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const existsR = await ps(`Test-Path "${iniPath}"`)
      if (existsR.out.trim().toLowerCase() !== 'true') { s('CitizenFX.ini not found.', 'warn'); return }
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      const content = contentR.out || ''
      const pairs = [['MaxLodDistance','30'],['MaxObjectLodDistance','30']]
      let c = content
      for (const [key, val] of pairs) {
        if (new RegExp(`${key}\\s*=`,'i').test(c)) {
          c = c.replace(new RegExp(`${key}\\s*=\\s*\\d+(\\.\\d+)?`,'i'), `${key}=${val}`)
        } else {
          c += `\n${key}=${val}`
        }
      }
      await ps(`Set-Content -Path "${iniPath}" -Value @'
${c}
'@ -Encoding ASCII -Force`)
      s('LOD distances reduced — fewer objects rendered at distance, meaningful FPS gain in busy areas.', 'ok')
    },
    restore: async (s, ps) => {
      const iniPath = `${process.env.LOCALAPPDATA}\\FiveM\\FiveM.app\\CitizenFX.ini`
      const contentR = await ps(`Get-Content "${iniPath}" -Raw -EA SilentlyContinue`)
      let c = contentR.out || ''
      c = c.replace(/MaxLodDistance\s*=\s*\d+(\.\d+)?/i, 'MaxLodDistance=100')
      c = c.replace(/MaxObjectLodDistance\s*=\s*\d+(\.\d+)?/i, 'MaxObjectLodDistance=100')
      await ps(`Set-Content -Path "${iniPath}" -Value @'
${c}
'@ -Encoding ASCII -Force`)
      s('LOD distances restored to default (100).', 'ok')
    }
  },

  // ── Fixes ─────────────────────────────────────────────────────────────────
  'fix-valorant-cfg': {
    apply: async (s, ps) => {
      await ps('Set-ProcessMitigation -Name vgc.exe -Enable CFG -ErrorAction SilentlyContinue')
      await ps('Set-ProcessMitigation -Name VALORANT-Win64-Shipping.exe -Enable CFG -ErrorAction SilentlyContinue')
      await ps(`
        $r = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\kernel'
        Set-ItemProperty -Path $r -Name DisableExceptionChainValidation -Value 0 -Force -EA SilentlyContinue
        Set-ItemProperty -Path $r -Name KernelSEHOPEnabled -Value 1 -Force -EA SilentlyContinue
      `)
      s('Valorant CFG fix applied — Control Flow Guard enabled. Reboot may be required.', 'ok')
    },
    restore: async (s) => s('CFG settings are system-level — reboot to verify state.', 'info')
  },
  'fix-valorant-hvci': {
    apply: async (s, ps) => {
      await ps(`
        $path = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceGuard\\Scenarios\\HypervisorEnforcedCodeIntegrity'
        New-Item -Path $path -Force -EA SilentlyContinue | Out-Null
        Set-ItemProperty -Path $path -Name Enabled -Value 1 -Force
      `)
      s('HVCI Memory Integrity enabled. Reboot required for Valorant.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`
        $path = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceGuard\\Scenarios\\HypervisorEnforcedCodeIntegrity'
        Set-ItemProperty -Path $path -Name Enabled -Value 0 -Force -EA SilentlyContinue
      `)
      s('HVCI disabled. Reboot required.', 'ok')
    }
  },
  'fix-nvidia-cp': {
    apply: async (s, ps) => {
      const r = await ps(`
        $pkg = Get-AppxPackage -Name 'NVIDIACorp.NVIDIAControlPanel' -AllUsers -EA SilentlyContinue
        if ($pkg) {
          $loc = $pkg.InstallLocation
          $exe = Join-Path $loc 'nvcplui.exe'
          if (Test-Path $exe) {
            Start-Process $exe
            Write-Output "LAUNCHED:$exe"
          } else { Write-Output 'NOT_FOUND' }
        } else { Write-Output 'NO_PKG' }
      `)
      if (r.out.startsWith('LAUNCHED')) {
        s('NVIDIA Control Panel launched from AppX package.', 'ok')
      } else {
        const r2 = await ps(`
          $exe = (Get-Command 'nvcplui.exe' -EA SilentlyContinue)?.Source
          if (-not $exe) { $exe = 'C:\\Program Files\\NVIDIA Corporation\\Control Panel Client\\nvcplui.exe' }
          if (Test-Path $exe) { Start-Process $exe; Write-Output "OK" } else { Write-Output "NOT_FOUND" }
        `)
        if (r2.out.trim() === 'OK') s('NVIDIA Control Panel launched.', 'ok')
        else s('NVIDIA Control Panel not found. Install the latest NVIDIA driver.', 'warn')
      }
    },
    restore: async (s) => s('No restore needed for this fix.', 'info')
  },
  'fix-wifi': {
    apply: async (s, ps, cmd) => {
      // ── 1. Restore all Wi-Fi dependent services to Automatic + start them ──
      s('Restoring Wi-Fi services…', 'info')
      for (const svc of ['nsi', 'Dhcp', 'RpcSs', 'Wlansvc', 'WinHttpAutoProxySvc', 'dot3svc', 'WwanSvc']) {
        await ps(`
          $svc = Get-Service '${svc}' -EA SilentlyContinue
          if ($svc) {
            Set-Service '${svc}' -StartupType Automatic -EA SilentlyContinue
            Start-Service '${svc}' -EA SilentlyContinue
          }
        `)
      }
      s('  ✓ Wi-Fi services set to Automatic', 'ok')

      // ── 2. Find the actual wireless adapter name (don't hardcode "Wi-Fi") ──
      const adapterR = await ps(`
        $a = Get-NetAdapter | Where-Object {
          $_.PhysicalMediaType -eq 'Native 802.11' -or
          $_.PhysicalMediaType -eq 'Wireless LAN' -or
          $_.InterfaceDescription -like '*Wireless*' -or
          $_.InterfaceDescription -like '*Wi-Fi*' -or
          $_.InterfaceDescription -like '*802.11*' -or
          $_.Name -like '*Wi-Fi*' -or
          $_.Name -like '*Wireless*' -or
          $_.Name -like '*WLAN*'
        } | Select-Object -First 1
        if ($a) { Write-Output $a.Name } else { Write-Output 'NONE' }
      `)
      const wifiName = adapterR.out.trim()

      if (wifiName && wifiName !== 'NONE') {
        s(`  Found Wi-Fi adapter: ${wifiName}`, 'info')

        // ── 3. Re-enable the adapter if it was disabled ──────────────────────
        await ps(`Enable-NetAdapter -Name '${wifiName}' -Confirm:$false -EA SilentlyContinue`)
        s('  ✓ Wi-Fi adapter enabled', 'ok')

        // ── 4. Enable WLAN autoconfig on the detected interface ──────────────
        await ps(`netsh wlan set autoconfig enabled=yes interface='${wifiName}'`)
        s('  ✓ WLAN autoconfig enabled', 'ok')

        // ── 5. Undo any offload tweaks applied to the Wi-Fi adapter ─────────
        // These are the same settings NIC tweaks apply — reset to driver defaults
        await ps(`
          $n = '${wifiName}'
          $props = @(
            'Energy Efficient Ethernet',
            '*EEE',
            'Flow Control',
            'Interrupt Moderation',
            'Large Send Offload v2 (IPv4)',
            'Large Send Offload v2 (IPv6)',
            'Receive Side Scaling',
            'TCP Checksum Offload (IPv4)',
            'TCP Checksum Offload (IPv6)',
            'UDP Checksum Offload (IPv4)',
            'UDP Checksum Offload (IPv6)',
            '*PMARPOffload',
            '*PMNSOffload',
            '*PowerSavingMode',
            'Auto Power Saver',
            'Power Saving Mode'
          )
          foreach ($p in $props) {
            try { Reset-NetAdapterAdvancedProperty -Name $n -DisplayName $p -EA SilentlyContinue } catch {}
          }
        `)
        s('  ✓ Wi-Fi adapter offload settings reset to driver defaults', 'ok')

        // ── 6. Disable power management on the Wi-Fi adapter ─────────────────
        // "Allow the computer to turn off this device to save power" causes random drops
        await ps(`
          $adapterName = '${wifiName}'
          $pnpDev = Get-PnpDevice | Where-Object { $_.FriendlyName -like "*$adapterName*" -or $_.FriendlyName -like "*Wireless*" -or $_.FriendlyName -like "*Wi-Fi*" } | Select-Object -First 1
          if ($pnpDev) {
            $devId = $pnpDev.InstanceId -replace '\\\\','\\\\'
            $regPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\$($pnpDev.InstanceId)\\Device Parameters\\Power Policy"
            New-Item -Path $regPath -Force -EA SilentlyContinue | Out-Null
            Set-ItemProperty -Path $regPath -Name 'IdlePowerState' -Value 0 -Force -EA SilentlyContinue
          }
          # Also via WMI power management
          $wmi = Get-WmiObject MSPower_DeviceEnable -Namespace root\\wmi -EA SilentlyContinue | Where-Object { $_.InstanceName -like "*$adapterName*" }
          if ($wmi) { $wmi.Enable = $false; $wmi.Put() | Out-Null }
        `, 10000)
        s('  ✓ Wi-Fi power management disabled', 'ok')

      } else {
        s('  ⚠ No wireless adapter detected — skipping adapter-specific steps', 'warn')
      }

      // ── 7. Reset Winsock and TCP/IP stack ────────────────────────────────────
      // This is the most common cause of "Wi-Fi connected but no internet"
      s('Resetting TCP/IP stack and Winsock…', 'info')
      await cmd('netsh winsock reset catalog')
      await cmd('netsh int ip reset')
      await cmd('netsh int ipv4 reset')
      await cmd('netsh int ipv6 reset')
      s('  ✓ Winsock and TCP/IP stack reset', 'ok')

      // ── 8. Flush DNS and release/renew IP ───────────────────────────────────
      await cmd('ipconfig /flushdns')
      await cmd('ipconfig /registerdns')
      s('  ✓ DNS cache flushed', 'ok')

      // ── 9. Reset firewall to allow Wi-Fi traffic ─────────────────────────────
      await ps(`netsh advfirewall set allprofiles state on`)
      await ps(`netsh advfirewall reset`)
      s('  ✓ Firewall reset to defaults', 'ok')

      s('✓ Wi-Fi fix complete. REBOOT REQUIRED for TCP/IP stack reset to take effect.', 'ok')
    },
    restore: async (s) => s('Wi-Fi fix does not need reverting.', 'info')
  },
  'fix-bluetooth': {
    apply: async (s, _, cmd) => {
      for (const svc of ['bthserv', 'BthAvctpSvc', 'BluetoothUserService']) {
        await cmd(`sc config ${svc} start= auto`)
        await cmd(`sc start ${svc}`)
      }
      s('Bluetooth services restored. Toggle Bluetooth off and back on to reconnect.', 'ok')
    },
    restore: async (s) => s('Bluetooth fix does not need reverting.', 'info')
  },
  'fix-2502': {
    apply: async (s, ps) => {
      await ps(`
        $tmp = 'C:\\Windows\\Temp'
        if (-not (Test-Path $tmp)) { New-Item -Path $tmp -ItemType Directory -Force | Out-Null }
        icacls 'C:\\Windows\\Temp' /grant 'Everyone:(F)' /T /Q 2>$null | Out-Null
        icacls 'C:\\Windows\\Temp' /grant 'SYSTEM:(F)' /T /Q 2>$null | Out-Null
        icacls "$env:TEMP" /grant 'Everyone:(F)' /T /Q 2>$null | Out-Null
      `)
      s('Error 2502/2503 fix applied — Temp directory permissions corrected.', 'ok')
    },
    restore: async (s) => s('Permission fix does not need reverting.', 'info')
  },
  'fix-ms-store': {
    apply: async (s, ps) => {
      s('Resetting Microsoft Store and Windows Update…', 'info')
      await ps('Get-AppxPackage -AllUsers -Name "Microsoft.WindowsStore" | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\\AppXManifest.xml" -EA SilentlyContinue }')
      await ps('wsreset.exe')
      for (const svc of ['wuauserv', 'BITS', 'cryptSvc', 'msiserver']) {
        await ps(`Start-Service -Name '${svc}' -EA SilentlyContinue`)
        await ps(`Set-Service -Name '${svc}' -StartupType Automatic -EA SilentlyContinue`)
      }
      s('Microsoft Store and Windows Update services restored.', 'ok')
    },
    restore: async (s) => s('No restore needed — services are back to default.', 'info')
  },
  'fix-audio': {
    apply: async (s, ps, _, cmd) => {
      s('Restarting audio services…', 'info')
      await ps('Stop-Service -Name AudioSrv,Audiosrv,AudioEndpointBuilder -Force -EA SilentlyContinue')
      await ps(`
        $dlls = @('dsound.dll','quartz.dll','wdmaud.drv','mmdevapi.dll')
        foreach ($d in $dlls) {
          $p = "C:\\Windows\\System32\\$d"
          if (Test-Path $p) { regsvr32 /s $p 2>$null }
        }
      `)
      await ps('Start-Service -Name AudioEndpointBuilder,AudioSrv -EA SilentlyContinue')
      s('Audio services restarted and DLLs re-registered. If no sound, reboot.', 'ok')
    },
    restore: async (s) => s('Audio fix does not need reverting.', 'info')
  },
  'fix-winsock': {
    apply: async (s, _, cmd) => {
      await cmd('netsh winsock reset catalog')
      await cmd('netsh int ip reset resetlog.txt')
      s('Winsock and TCP/IP stack reset. REBOOT REQUIRED to take effect.', 'ok')
    },
    restore: async (s) => s('Winsock reset is permanent — reboot was applied.', 'info')
  },
  'reboot-now': {
    apply: async (s, _, cmd) => {
      s('Rebooting in 5 seconds…', 'info')
      await cmd('shutdown /r /t 5 /c "Jylli Tool: reboot to complete Wi-Fi fix"')
    },
    restore: async () => {}
  },
  'fix-windows-defender': {
    apply: async (s, ps, _, cmd) => {
      s('Restoring Windows Defender…', 'info')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Name DisableAntiSpyware -Value 0 -Force -EA SilentlyContinue')
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Name DisableAntiSpyware -EA SilentlyContinue')
      await cmd('sc config WinDefend start= auto')
      await cmd('sc start WinDefend')
      await ps('Set-MpPreference -DisableRealtimeMonitoring $false -EA SilentlyContinue')
      await ps('Start-MpScan -ScanType QuickScan -EA SilentlyContinue')
      s('Windows Defender restored and re-enabled.', 'ok')
    },
    restore: async (s) => s('Defender restore fix does not need reverting.', 'info')
  },
  'fix-print-spooler': {
    apply: async (s, ps, _, cmd) => {
      await cmd('net stop Spooler')
      await ps('Remove-Item -Path "C:\\Windows\\System32\\spool\\PRINTERS\\*" -Recurse -Force -EA SilentlyContinue')
      await cmd('net start Spooler')
      s('Print spooler cleared and restarted — stuck print jobs removed.', 'ok')
    },
    restore: async (s) => s('Print spooler fix does not need reverting.', 'info')
  },
  'fix-search': {
    apply: async (s, ps, _, cmd) => {
      await cmd('sc stop WSearch')
      await ps(`
        $dbPath = "$env:ProgramData\\Microsoft\\Search\\Data\\Applications\\Windows\\Windows.edb"
        if (Test-Path $dbPath) {
          try { Remove-Item $dbPath -Force -EA SilentlyContinue } catch {}
        }
      `)
      await cmd('sc start WSearch')
      await ps('Update-Help -EA SilentlyContinue')
      s('Windows Search index rebuilt. Indexing will continue in the background — it may take a few minutes.', 'ok')
    },
    restore: async (s) => s('Search fix does not need reverting.', 'info')
  },

  // ── New General Tweaks (from screenshots) ─────────────────────────────────
  'raw-aim-curve': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseSpeed -Value "0" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold1 -Value "0" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold2 -Value "0" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name SmoothMouseXCurve -Value ([byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xC0,0xCC,0x0C,0x00,0x00,0x00,0x00,0x00,0x80,0x99,0x19,0x00,0x00,0x00,0x00,0x00,0x40,0x66,0x26,0x00,0x00,0x00,0x00,0x00,0x00,0x33,0x33,0x00,0x00,0x00,0x00,0x00)) -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name SmoothMouseYCurve -Value ([byte[]](0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x38,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xA8,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xE0,0x00,0x00,0x00,0x00,0x00)) -Force -EA SilentlyContinue')
      s('Raw Aim Curve applied — 0/0 pointer thresholds, zero accel, optimised X/Y curves.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseSpeed -Value "1" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold1 -Value "6" -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Mouse" -Name MouseThreshold2 -Value "10" -Force')
      s('Mouse curve restored to Windows defaults.', 'ok')
    }
  },
  'proximity-deny': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\proximity" -Name Value -Value "Deny" -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\bluetoothSync" -Name Value -Value "Deny" -Force -EA SilentlyContinue')
      s('Proximity/device coupling denied — stops auto-discovery chatter on idle.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\proximity" -Name Value -Value "Allow" -Force -EA SilentlyContinue')
      s('Proximity access restored.', 'ok')
    }
  },
  'reality-telephony-off': {
    apply: async (s, ps, _, cmd) => {
      await cmd('sc config MixedRealityOpenXRSvc start= disabled'); await cmd('sc stop MixedRealityOpenXRSvc')
      await cmd('sc config TapiSrv start= disabled'); await cmd('sc stop TapiSrv')
      await cmd('sc config PhoneSvc start= disabled'); await cmd('sc stop PhoneSvc')
      s('Mixed Reality and Telephony API services disabled.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config TapiSrv start= manual'); await cmd('sc start TapiSrv')
      s('Telephony service restored.', 'ok')
    }
  },
  'cloud-sync-off': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" -Name DisableSettingSync -Value 2 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" -Name DisableSettingSyncUserOverride -Value 1 -Force')
      s('Windows cloud settings sync disabled — prevents background sync and conflict scans.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" -Name DisableSettingSync -EA SilentlyContinue')
      s('Cloud sync restored.', 'ok')
    }
  },
  'do-solo-mode': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Name DODownloadMode -Value 1 -Force')
      s('Delivery Optimization set to Microsoft-only (no peer-to-peer). Prevents LAN/WAN spikes.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization" -Name DODownloadMode -EA SilentlyContinue')
      s('Delivery Optimization restored to default.', 'ok')
    }
  },
  'nudge-blocker': {
    apply: async (s, ps) => {
      const base = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager'
      const keys = ['SoftLandingEnabled','SubscribedContent-338389Enabled','SubscribedContent-338388Enabled',
        'SubscribedContent-310093Enabled','RotatingLockScreenEnabled','RotatingLockScreenOverlayEnabled',
        'SystemPaneSuggestionsEnabled','OemPreInstalledAppsEnabled','PreInstalledAppsEnabled',
        'SilentInstalledAppsEnabled','ContentDeliveryAllowed']
      for (const k of keys) await ps(`Set-ItemProperty -Path "${base}" -Name "${k}" -Value 0 -Force -EA SilentlyContinue`)
      s('Account nudges, feedback flags, and content suggestions disabled.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager'
      await ps(`Set-ItemProperty -Path "${base}" -Name SoftLandingEnabled -Value 1 -Force -EA SilentlyContinue`)
      s('Nudge blocker restored.', 'ok')
    }
  },
  'rt-scan-trim': {
    apply: async (s, ps) => {
      await ps('Set-MpPreference -SubmitSamplesConsent 2 -EA SilentlyContinue')
      await ps('Set-MpPreference -MAPSReporting 0 -EA SilentlyContinue')
      await ps('Set-MpPreference -DisableIOAVProtection $false -EA SilentlyContinue')
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender\\Spynet" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender\\Spynet" -Name SpynetReporting -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender\\Spynet" -Name SubmitSamplesConsent -Value 2 -Force')
      s('Defender cloud reporting lowered — fewer real-time IO inspections. Reboot recommended.', 'warn')
    },
    restore: async (s, ps) => {
      await ps('Set-MpPreference -SubmitSamplesConsent 1 -EA SilentlyContinue')
      await ps('Set-MpPreference -MAPSReporting 2 -EA SilentlyContinue')
      s('Defender reporting restored.', 'ok')
    }
  },
  'search-box-local': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name DisableWebSearch -Value 1 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name ConnectedSearchUseWeb -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Search" -Name BingSearchEnabled -Value 0 -Force -EA SilentlyContinue')
      s('Search box made local-only — Bing search and suggestions disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name DisableWebSearch,ConnectedSearchUseWeb -EA SilentlyContinue')
      s('Search box restored to default.', 'ok')
    }
  },
  'search-cloud-off': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCloudSearch -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Search" -Name CortanaConsent -Value 0 -Force -EA SilentlyContinue')
      s('AAD/MSA cloud search and web usage disabled. Less network jitter on Start open.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCloudSearch -EA SilentlyContinue')
      s('Cloud search restored.', 'ok')
    }
  },
  'security-quiet': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender Security Center\\Notifications" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender Security Center\\Notifications" -Name DisableNotifications -Value 1 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender Security Center\\Notifications" -Name DisableEnhancedNotifications -Value 1 -Force')
      s('Defender Security Center alerts and lock-screen toasts suppressed.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender Security Center\\Notifications" -Name DisableNotifications -EA SilentlyContinue')
      s('Security notifications restored.', 'ok')
    }
  },
  'security-toast-off': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.SystemToast.SecurityAndMaintenance" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.SystemToast.SecurityAndMaintenance" -Name Enabled -Value 0 -Force')
      await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.Defender.SecurityCenter" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.Defender.SecurityCenter" -Name Enabled -Value 0 -Force')
      s('Defender and Security Center toast notifications disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.Defender.SecurityCenter" -Name Enabled -EA SilentlyContinue')
      s('Security toast notifications restored.', 'ok')
    }
  },
  'sensor-suite-off': {
    apply: async (s, ps, _, cmd) => {
      for (const svc of ['SensorService','SensrSvc','SensorDataService','lfsvc']) {
        await cmd(`sc config ${svc} start= disabled`); await cmd(`sc stop ${svc}`)
      }
      s('Sensor, location, and awareness services disabled. More stable CPU clocks while gaming.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config SensorService start= manual'); await cmd('sc start SensorService')
      s('Sensor services restored.', 'ok')
    }
  },
  'sticky-keys-guard': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Accessibility\\StickyKeys" -Name Flags -Value "506" -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Accessibility\\ToggleKeys" -Name Flags -Value "58" -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Accessibility\\Keyboard Response" -Name Flags -Value "122" -Force -EA SilentlyContinue')
      s('StickyKeys popup disabled — won\'t interrupt tournaments with focus-steal dialogs.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKCU:\\Control Panel\\Accessibility\\StickyKeys" -Name Flags -Value "510" -Force -EA SilentlyContinue')
      s('StickyKeys defaults restored.', 'ok')
    }
  },
  'telemetry-zero': {
    apply: async (s, ps, _, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" -Name AllowTelemetry -Value 0 -Force')
      await cmd('sc config DiagTrack start= disabled'); await cmd('sc stop DiagTrack')
      await cmd('sc config dmwappushservice start= disabled'); await cmd('sc stop dmwappushservice')
      await cmd('sc config diagnosticshub.standardcollector.service start= disabled')
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Diagnostics\\DiagTrack" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Diagnostics\\DiagTrack" -Name DiagTrackAuthorization -Value 0 -Force')
      s('Telemetry forced to 0 — feedback prompts hidden, diagnostic pipelines cut. Reboot recommended.', 'warn')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config DiagTrack start= auto'); await cmd('sc start DiagTrack')
      s('Telemetry services restored.', 'ok')
    }
  },
  'usb-power-guard': {
    apply: async (s, ps, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setdcvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0')
      await cmd('powercfg /setactive SCHEME_CURRENT')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\usb" -Name DisableSelectiveSuspend -Value 1 -Force -EA SilentlyContinue')
      s('USB Power Guard applied — selective suspend off, root hub idle disabled. Eliminates input lag spikes.', 'ok')
    },
    restore: async (s, ps, _, cmd) => {
      await cmd('powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 1')
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\usb" -Name DisableSelectiveSuspend -EA SilentlyContinue')
      s('USB power management restored.', 'ok')
    }
  },
  'biometrics-off': {
    apply: async (s, ps, _, cmd) => {
      await cmd('sc config WbioSrvc start= disabled'); await cmd('sc stop WbioSrvc')
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Biometrics" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Biometrics" -Name Enabled -Value 0 -Force')
      s('Windows Hello and biometric stack disabled. Removes lock-screen camera overhead.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config WbioSrvc start= auto'); await cmd('sc start WbioSrvc')
      s('Biometrics service restored.', 'ok')
    }
  },
  'bluetooth-hard-off': {
    apply: async (s, ps, _, cmd) => {
      for (const svc of ['bthserv','BthAvctpSvc','BluetoothUserService','BTAGService']) {
        await cmd(`sc config ${svc} start= disabled`); await cmd(`sc stop ${svc}`)
      }
      s('Bluetooth core and companion services fully disabled. BT stack cannot wake CPU.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config bthserv start= auto'); await cmd('sc start bthserv')
      await cmd('sc config BthAvctpSvc start= auto'); await cmd('sc start BthAvctpSvc')
      s('Bluetooth services restored.', 'ok')
    }
  },
  'boot-sound-off': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Authentication\\LogonUI\\BootAnimation" -Name DisableStartupSound -Value 1 -Force -EA SilentlyContinue')
      await ps('Set-ItemProperty -Path "HKCU:\\AppEvents\\Schemes\\Apps\\.Default\\WindowsLogon\\.Current" -Name "(Default)" -Value "" -Force -EA SilentlyContinue')
      s('Boot startup sound disabled in two locations.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Authentication\\LogonUI\\BootAnimation" -Name DisableStartupSound -Value 0 -Force -EA SilentlyContinue')
      s('Boot sound restored.', 'ok')
    }
  },
  'compat-scanner-off': {
    apply: async (s, ps) => {
      await ps('Disable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "ProgramDataUpdater" -EA SilentlyContinue')
      await ps('Disable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "AitAgent" -EA SilentlyContinue')
      await ps('Disable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "Microsoft Compatibility Appraiser" -EA SilentlyContinue')
      s('AppCompat inventory and UAR analytics tasks disabled. Prevents catalog scans that wake disks and CPU.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Enable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "ProgramDataUpdater" -EA SilentlyContinue')
      s('AppCompat tasks re-enabled.', 'ok')
    }
  },
  'compat-tasks-off': {
    apply: async (s, ps) => {
      for (const task of ['ProgramDataUpdater','AitAgent','StartupAppTask']) {
        await ps(`Disable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "${task}" -EA SilentlyContinue`)
      }
      await ps('Disable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\DiskDiagnostic\\" -TaskName "Microsoft-Windows-DiskDiagnosticDataCollector" -EA SilentlyContinue')
      s('Application Experience telemetry tasks disabled. Smoother frametime plots.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Enable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\Application Experience\\" -TaskName "StartupAppTask" -EA SilentlyContinue')
      s('App experience tasks re-enabled.', 'ok')
    }
  },
  'cortana-blackout': {
    apply: async (s, ps) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCortana -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCortanaAboveLock -Value 0 -Force')
      await ps('Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Search" -Name CortanaEnabled -Value 0 -Force -EA SilentlyContinue')
      s('Cortana/Cloud web features disabled — search won\'t ping the web while typing.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" -Name AllowCortana,AllowCortanaAboveLock -EA SilentlyContinue')
      s('Cortana restored.', 'ok')
    }
  },
  'cortana-remnant-off': {
    apply: async (s, ps) => {
      const r = await ps('Get-AppxPackage -AllUsers -Name "Microsoft.549981C3F5F10" -EA SilentlyContinue | Remove-AppxPackage -AllUsers -EA SilentlyContinue; Write-Output "done"')
      s('Cortana app package (Microsoft.549981C3F5F10) removal attempted. Reboot to clear remnants.', r.ok ? 'ok' : 'warn')
    },
    restore: async (s) => s('Cannot restore removed AppX package. Reinstall from Microsoft Store.', 'info')
  },
  'cpu-focus-bias': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" -Name Win32PrioritySeparation -Value 38 -Force')
      s('Win32PrioritySeparation=38 (hex) — foreground quantum boost maximised. Scheduler favours active game.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" -Name Win32PrioritySeparation -Value 2 -Force')
      s('Priority separation restored to default (2).', 'ok')
    }
  },
  'cursor-max-rate': {
    apply: async (s, ps) => {
      await ps(`
        $path = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\mouclass\\Parameters'
        New-Item -Path $path -Force -EA SilentlyContinue | Out-Null
        Set-ItemProperty -Path $path -Name MouseDataQueueSize -Value 20 -Force
      `)
      s('Mouse input queue size increased — buffers more samples for high-polling mice.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\mouclass\\Parameters" -Name MouseDataQueueSize -EA SilentlyContinue')
      s('Mouse class parameters restored.', 'ok')
    }
  },
  'defender-core-off': {
    apply: async (s, ps, _, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Name DisableAntiSpyware -Value 1 -Force')
      await ps('Set-MpPreference -DisableRealtimeMonitoring $true -EA SilentlyContinue')
      await cmd('sc config WinDefend start= disabled')
      s('Defender core real-time protection disabled. SECURITY RISK — use only offline/tournament PCs. Reboot required.', 'warn')
    },
    restore: async (s, ps, _, cmd) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Name DisableAntiSpyware -Value 0 -Force -EA SilentlyContinue')
      await ps('Remove-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender" -Name DisableAntiSpyware -EA SilentlyContinue')
      await cmd('sc config WinDefend start= auto'); await cmd('sc start WinDefend')
      await ps('Set-MpPreference -DisableRealtimeMonitoring $false -EA SilentlyContinue')
      s('Windows Defender restored and re-enabled.', 'ok')
    }
  },
  'defender-quiet': {
    apply: async (s, ps) => {
      await ps('Set-MpPreference -MAPSReporting 0 -EA SilentlyContinue')
      await ps('Set-MpPreference -SubmitSamplesConsent 2 -EA SilentlyContinue')
      s('Defender cloud reporting level set to 0 (SpyNetReporting). Background telemetry during scans reduced.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-MpPreference -MAPSReporting 2 -EA SilentlyContinue')
      await ps('Set-MpPreference -SubmitSamplesConsent 1 -EA SilentlyContinue')
      s('Defender reporting restored.', 'ok')
    }
  },
  'diag-tasks-off': {
    apply: async (s, ps) => {
      const tasks = [
        ['\\Microsoft\\Windows\\Diagnosis\\', 'Scheduled'],
        ['\\Microsoft\\Windows\\DiskFootprint\\', 'Diagnostics'],
        ['\\Microsoft\\Windows\\DiskDiagnostic\\', 'Microsoft-Windows-DiskDiagnosticDataCollector'],
        ['\\Microsoft\\Windows\\WDI\\', 'ResolutionHost'],
        ['\\Microsoft\\Windows\\Diagnosis\\', 'RecommendedTroubleshootingScanner'],
      ]
      for (const [path, name] of tasks) {
        await ps(`Disable-ScheduledTask -TaskPath "${path}" -TaskName "${name}" -EA SilentlyContinue`)
      }
      s('Power diagnostics, error reporting, disk footprint, and scheduler diagnosis tasks disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Enable-ScheduledTask -TaskPath "\\Microsoft\\Windows\\DiskDiagnostic\\" -TaskName "Microsoft-Windows-DiskDiagnosticDataCollector" -EA SilentlyContinue')
      s('Diagnostic tasks re-enabled.', 'ok')
    }
  },
  'diagnostics-off': {
    apply: async (s, ps, _, cmd) => {
      await ps('New-Item -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WDI" -Force -EA SilentlyContinue | Out-Null')
      await ps('Set-ItemProperty -Path "HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WDI" -Name ScenarioExecutionEnabled -Value 0 -Force')
      await cmd('sc config WdiServiceHost start= disabled')
      await cmd('sc config WdiSystemHost start= disabled')
      s('Diagnostics policy and host services disabled. Less background scheduler interference.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config WdiServiceHost start= demand')
      await cmd('sc config WdiSystemHost start= demand')
      s('Diagnostics services restored.', 'ok')
    }
  },

  // ── New Advanced Tweaks (from screenshots) ───────────────────────────────
  'apex-power-plan': {
    apply: async (s, _, cmd) => {
      // Duplicate Ultimate Performance plan and apply extreme settings
      const r = await cmd('powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61')
      const m = r.out.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i)
      if (!m) { s('Could not create power plan.', 'warn'); return }
      const guid = m[1]
      await cmd(`powercfg -setactive ${guid}`)
      await cmd(`powercfg /setacvalueindex ${guid} SUB_PROCESSOR CPMINCORES 100`)
      await cmd(`powercfg /setacvalueindex ${guid} SUB_PROCESSOR PROCTHROTTLEMIN 100`)
      await cmd(`powercfg /setacvalueindex ${guid} SUB_PROCESSOR PROCTHROTTLEMAX 100`)
      await cmd(`powercfg /setacvalueindex ${guid} SUB_PROCESSOR IDLEPROMOTE 0`)
      await cmd(`powercfg /setacvalueindex ${guid} 501a4d13-42af-4429-9ac1-df54c6bf3fc2 ee12f906-d277-404b-b6da-e5fa1a576df5 0`)
      await cmd(`powercfg /setacvalueindex ${guid} 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0`)
      await cmd(`powercfg -setactive ${guid}`)
      s(`Apex Power Plan activated (${guid}) — maximum CPU, PCIe and USB performance. Reboot recommended.`, 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('powercfg -setactive 381b4222-f694-41f0-9685-ff5bb260df2e')
      s('Balanced power plan restored.', 'ok')
    }
  },
  'gpu-hags-off': {
    apply: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 1 -Force')
      await ps('New-Item -Path "HKCU:\\Software\\Microsoft\\DirectX\\UserGpuPreferences" -Force -EA SilentlyContinue | Out-Null')
      s('Hardware-Accelerated GPU Scheduling (HAGS) disabled. Reboot required.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name HwSchMode -Value 2 -Force')
      s('HAGS re-enabled. Reboot required.', 'ok')
    }
  },
  'hvci-offload': {
    apply: async (s, ps) => {
      await ps(`
        $path = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceGuard\\Scenarios\\HypervisorEnforcedCodeIntegrity'
        New-Item -Path $path -Force -EA SilentlyContinue | Out-Null
        Set-ItemProperty -Path $path -Name Enabled -Value 0 -Force
        Set-ItemProperty -Path $path -Name WasEnabledBy -Value 0 -Force
      `)
      s('HVCI (Memory Integrity) disabled. Removes virtualization overhead. SECURITY RISK. Reboot required.', 'warn')
    },
    restore: async (s, ps) => {
      await ps(`
        $path = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceGuard\\Scenarios\\HypervisorEnforcedCodeIntegrity'
        Set-ItemProperty -Path $path -Name Enabled -Value 1 -Force -EA SilentlyContinue
      `)
      s('HVCI re-enabled. Reboot required.', 'ok')
    }
  },
  'memory-guard-tune': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management'
      await ps(`Set-ItemProperty -Path "${base}" -Name DisablePagingExecutive -Value 1 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name LargeSystemCache -Value 0 -Force`)
      await ps(`Set-ItemProperty -Path "${base}" -Name ClearPageFileAtShutdown -Value 0 -Force -EA SilentlyContinue`)
      await ps(`Set-ItemProperty -Path "${base}" -Name PagingFiles -Value "C:\\pagefile.sys 0 0" -Force -EA SilentlyContinue`)
      s('Memory Guard Tune applied — paging executive disabled, large cache off. Reboot required.', 'warn')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management'
      await ps(`Set-ItemProperty -Path "${base}" -Name DisablePagingExecutive -Value 0 -Force`)
      s('Memory Management restored.', 'ok')
    }
  },
  'nvidia-pstate-lock': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name DisableDynamicPstates -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerLevel -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PerfLevelSrc -Value 8738 -Force -EA SilentlyContinue`)
      }
      s('NVIDIA P-State locked — dynamic P-states disabled. GPU stays at max clocks. Reboot required.', 'warn')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name DisableDynamicPstates -EA SilentlyContinue`)
        await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name PerfLevelSrc -EA SilentlyContinue`)
      }
      s('NVIDIA P-States restored to dynamic.', 'ok')
    }
  },
  'full-mitigation-wipe': {
    apply: async (s, ps) => {
      await ps(`
        Set-ProcessMitigation -System -Disable CFG,StrictCFG,SEHOP,ForceRelocateImages,RequireInfo,NullPage -ErrorAction SilentlyContinue
      `)
      await ps(`
        $base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management'
        Set-ItemProperty -Path $base -Name FeatureSettingsOverride -Value 3 -Force
        Set-ItemProperty -Path $base -Name FeatureSettingsOverrideMask -Value 3 -Force
      `)
      s('Full Mitigation Wipe applied — all Windows process mitigations disabled at system level. MAJOR SECURITY RISK. Reboot required.', 'warn')
    },
    restore: async (s, ps) => {
      await ps(`
        Remove-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management' -Name FeatureSettingsOverride -EA SilentlyContinue
        Remove-ItemProperty -Path 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management' -Name FeatureSettingsOverrideMask -EA SilentlyContinue
      `)
      s('CPU mitigations restored. Reboot required.', 'ok')
    }
  },
  'nvidia-profile': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerEnable -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerLevel -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PowerMizerLevelAC -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name PerfLevelSrc -Value 8738 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name D3PCLatency -Value 1 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name F1MaxLatency -Value 0 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name RMAERQSizeMultiplier -Value 32 -Force -EA SilentlyContinue`)
      }
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrLevel -Value 3 -Force')
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" -Name TdrDelay -Value 10 -Force')
      s('NVIDIA Profile applied — max perf mode, ultra low latency, DMA power gating disabled. Reboot required.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const name of ['PowerMizerEnable','PowerMizerLevel','D3PCLatency','F1MaxLatency','RMAERQSizeMultiplier']) {
        for (const idx of ['0000','0001']) await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name ${name} -EA SilentlyContinue`)
      }
      s('NVIDIA profile restored.', 'ok')
    }
  },
  'process-count-reduction': {
    apply: async (s, ps) => {
      const ramGB = require('os').totalmem() / (1024**3)
      const ramMB = Math.floor(ramGB * 1024)
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control'
      await ps(`Set-ItemProperty -Path "${base}\\Session Manager\\Memory Management" -Name SvcHostSplitThresholdInKB -Value ${ramMB * 1024} -Force`)
      s(`Process Count Reduction applied — SvcHost split threshold set to ${ramMB}MB. Background svchost processes will consolidate.`, 'ok')
    },
    restore: async (s, ps) => {
      await ps('Set-ItemProperty -Path "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" -Name SvcHostSplitThresholdInKB -Value 3670016 -Force')
      s('SvcHost split threshold restored.', 'ok')
    }
  },
  'amd-chill-off': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableChill -Value 0 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableBoost -Value 0 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableAntiLag -Value 0 -Force -EA SilentlyContinue`)
      }
      s('AMD Chill, Boost, and Anti-Lag disabled at driver class level. Frame pacing more predictable.', 'ok')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableChill -EA SilentlyContinue`)
        await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableBoost -EA SilentlyContinue`)
      }
      s('AMD Chill/Boost restored.', 'ok')
    }
  },
  'amd-power-hold': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) {
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableABC -Value 0 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name KMD_FRTEnabled -Value 0 -Force -EA SilentlyContinue`)
        await ps(`Set-ItemProperty -Path "${base}\\${idx}" -Name DalPowerPolicy -Value 0 -Force -EA SilentlyContinue`)
      }
      s('AMD Power Hold applied — power auto-enable, ULPS, and DMA power gating stopped. GPU clocks steadier.', 'warn')
    },
    restore: async (s, ps) => {
      const base = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}'
      for (const idx of ['0000','0001']) await ps(`Remove-ItemProperty -Path "${base}\\${idx}" -Name KMD_EnableABC,KMD_FRTEnabled -EA SilentlyContinue`)
      s('AMD power settings restored.', 'ok')
    }
  },
  'amd-service-trim': {
    apply: async (s, ps, _, cmd) => {
      for (const svc of ['AMD Crash Defender','AMD External Events Utility','AMD Log Utility']) {
        await cmd(`sc config "${svc}" start= disabled`)
        await cmd(`sc stop "${svc}"`)
      }
      await cmd('sc config amdfendr start= disabled'); await cmd('sc stop amdfendr')
      await cmd('sc config amdxc64 start= disabled')
      s('AMD Crash Defender and logging services disabled. Less DPC jitter and background CPU.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config "AMD External Events Utility" start= auto')
      await cmd('sc start "AMD External Events Utility"')
      s('AMD services restored.', 'ok')
    }
  },
  'corsair-audio-off': {
    apply: async (s, ps, _, cmd) => {
      await cmd('sc config CorsairVBusDriver start= disabled')
      await cmd('sc config CorsairGamingAudioConfig start= disabled')
      await cmd('sc stop CorsairGamingAudioConfig')
      s('Corsair audio config service disabled. Removes vendor background helper overhead.', 'ok')
    },
    restore: async (s, _, cmd) => {
      await cmd('sc config CorsairGamingAudioConfig start= auto')
      await cmd('sc start CorsairGamingAudioConfig')
      s('Corsair audio service restored.', 'ok')
    }
  },
  'capability-ruleset': {
    apply: async (s, ps) => {
      await ps(`
        $base = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Privacy'
        New-Item -Path $base -Force -EA SilentlyContinue | Out-Null
        Set-ItemProperty -Path $base -Name TailoredExperiencesWithDiagnosticDataEnabled -Value 0 -Force
        $base2 = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager'
        Set-ItemProperty -Path $base2 -Name SchemaVersion -Value 3 -Force -EA SilentlyContinue
      `)
      s('Capability Ruleset applied — default app capability access configured, personal data listeners reduced.', 'ok')
    },
    restore: async (s, ps) => {
      await ps('Remove-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Privacy" -Name TailoredExperiencesWithDiagnosticDataEnabled -EA SilentlyContinue')
      s('Capability ruleset restored.', 'ok')
    }
  },
}

// ─── Fix IPC handler ──────────────────────────────────────────────────────────
ipcMain.handle('run-fix', async (_, fixId) => {
  const send = (msg, level = 'info') => {
    mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  }
  send(`Running fix: ${fixId}`, 'head')
  try {
    await TWEAKS[fixId]?.apply?.(send, runPS, runCmd, regAdd, regDelete)
    return { ok: true }
  } catch (e) {
    send(`Fix error: ${e.message}`, 'err')
    return { ok: false, error: e.message }
  }
})

// ─── Windows Health Check (SFC + DISM + disk) ────────────────────────────────
ipcMain.handle('run-windows-health', async () => {
  const send = (msg, level = 'info', pct = null) =>
    mainWindow?.webContents.send('windows-health-progress', { msg, level, pct })

  setDiscordOverride(HEALTH_PHRASES, 'Checking DISM…')
  const results = { sfc: null, dism: null, disk: null, ok: true, issues: [] }

  // 1. DISM /CheckHealth (fast, ~5s — just reads the component store flag)
  send('Running DISM /CheckHealth…', 'info', 5)
  const dismCheck = await runPS('DISM /Online /Cleanup-Image /CheckHealth 2>&1')
  if (/no component store corruption detected/i.test(dismCheck.out)) {
    results.dism = 'healthy'
    send('✓ DISM: Component store is clean', 'ok', 25)
  } else if (/component store is repairable/i.test(dismCheck.out)) {
    results.dism = 'repairable'
    results.issues.push('Component store needs repair (DISM /RestoreHealth)')
    send('⚠ DISM: Component store is repairable — run DISM /RestoreHealth to fix', 'warn', 25)
  } else if (/component store is corrupted/i.test(dismCheck.out)) {
    results.dism = 'corrupted'
    results.ok = false
    results.issues.push('Component store is corrupted — DISM /RestoreHealth required')
    send('✗ DISM: Component store corrupted — repair recommended', 'err', 25)
  } else {
    results.dism = 'unknown'
    send('→ DISM: Could not determine status', 'info', 25)
  }

  // 2. SFC /verifyonly (read-only scan, ~60-120s — does not modify files)
  updateDiscordOverrideState('Running SFC scan…')
  send('Running SFC /verifyonly — scanning protected system files… (this takes 1–2 min)', 'info', 30)
  const sfcOut = await runPS('sfc /verifyonly 2>&1', 180000)
  const sfcText = sfcOut.out || ''
  if (/did not find any integrity violations/i.test(sfcText) || /Windows Resource Protection did not find any/i.test(sfcText)) {
    results.sfc = 'clean'
    send('✓ SFC: No integrity violations found', 'ok', 70)
  } else if (/found corrupt files/i.test(sfcText) || /found integrity violations/i.test(sfcText)) {
    results.sfc = 'corrupt'
    results.ok = false
    results.issues.push('SFC found corrupt system files — run SFC /scannow to repair')
    send('✗ SFC: Corrupt system files detected — SFC /scannow repair recommended', 'err', 70)
  } else if (/could not perform the requested operation/i.test(sfcText)) {
    results.sfc = 'pending'
    results.issues.push('SFC pending reboot — a previous repair is awaiting restart')
    send('⚠ SFC: Pending reboot from previous repair', 'warn', 70)
  } else {
    results.sfc = 'unknown'
    send('→ SFC: Scan completed (status unclear — likely clean)', 'info', 70)
  }

  // 3. Quick disk health (SMART status via WMI — instant)
  updateDiscordOverrideState('Checking disk SMART…')
  send('Checking disk SMART status…', 'info', 75)
  const diskOut = await runPS(`Get-WmiObject -Namespace root\\wmi -Class MSStorageDriver_FailurePredictStatus -EA SilentlyContinue | Select-Object InstanceName,PredictFailure | ConvertTo-Json -Compress`)
  let diskIssue = false
  try {
    const parsed = JSON.parse(diskOut.out.trim())
    const drives = Array.isArray(parsed) ? parsed : [parsed]
    diskIssue = drives.some(d => d?.PredictFailure === true)
  } catch {}
  if (diskIssue) {
    results.disk = 'warning'
    results.ok = false
    results.issues.push('SMART reports a disk failure predicted — back up data immediately')
    send('✗ Disk: SMART failure predicted — back up your data immediately!', 'err', 90)
  } else {
    results.disk = 'ok'
    send('✓ Disk: SMART status OK', 'ok', 90)
  }

  send(results.ok ? '✓ Windows Health check complete — no issues found' : `⚠ Health check complete — ${results.issues.length} issue(s) found`, results.ok ? 'ok' : 'warn', 100)
  clearDiscordOverride()
  webhookHealthCheck(results)
  return results
})

// ─── Pre-Flight Diagnostic ────────────────────────────────────────────────────
ipcMain.handle('run-preflight-scan', async () => {
  setDiscordOverride(SCAN_PHRASES, 'Scanning system…')
  const detections = []
  let score = 100

  // ── 1. Custom OS detection ───────────────────────────────────────────────
  const oemR = await runPS(`
    $oemPath = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\OEMInformation'
    $ntPath  = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion'
    $man  = (Get-ItemProperty $oemPath -EA SilentlyContinue).Manufacturer
    $mod  = (Get-ItemProperty $oemPath -EA SilentlyContinue).Model
    $url  = (Get-ItemProperty $oemPath -EA SilentlyContinue).SupportURL
    $org  = (Get-ItemProperty $ntPath  -EA SilentlyContinue).RegisteredOrganization
    $own  = (Get-ItemProperty $ntPath  -EA SilentlyContinue).RegisteredOwner
    Write-Output "MAN=$man"
    Write-Output "MOD=$mod"
    Write-Output "URL=$url"
    Write-Output "ORG=$org"
    Write-Output "OWN=$own"
    Write-Output "AME=$(Test-Path 'C:\\ProgramData\\AME' -EA SilentlyContinue)"
    Write-Output "AMEWIZ=$(Test-Path 'C:\\ProgramData\\AME\\ame.conf' -EA SilentlyContinue)"
    Write-Output "ATLAS_MOD=$(Test-Path 'C:\\Windows\\AtlasModules' -EA SilentlyContinue)"
    Write-Output "ATLAS_DIR=$(Test-Path 'C:\\AtlasOS' -EA SilentlyContinue)"
    Write-Output "REVI_DIR=$(Test-Path 'C:\\ReviOS' -EA SilentlyContinue)"
    Write-Output "NSUDO=$(Test-Path 'C:\\Windows\\System32\\NSudo.exe' -EA SilentlyContinue)"
    Write-Output "PLAYBOOKS=$(if (Test-Path 'C:\\ProgramData\\AME\\Playbooks' -EA SilentlyContinue) { (Get-ChildItem 'C:\\ProgramData\\AME\\Playbooks' -EA SilentlyContinue | Select-Object -First 1).Name } else { '' })"
  `)
  const oemMan  = oemR.out.match(/MAN=(.*)/)?.[1]?.trim() || ''
  const oemMod  = oemR.out.match(/MOD=(.*)/)?.[1]?.trim() || ''
  const oemUrl  = oemR.out.match(/URL=(.*)/)?.[1]?.trim() || ''
  const regOrg  = oemR.out.match(/ORG=(.*)/)?.[1]?.trim() || ''
  const regOwn  = oemR.out.match(/OWN=(.*)/)?.[1]?.trim() || ''
  const hasAME  = /AME=True/i.test(oemR.out)
  const hasAMEConf = /AMEWIZ=True/i.test(oemR.out)
  const hasAtlasMod = /ATLAS_MOD=True/i.test(oemR.out)
  const hasAtlasDir = /ATLAS_DIR=True/i.test(oemR.out)
  const hasReviDir  = /REVI_DIR=True/i.test(oemR.out)
  const hasNSudo    = /NSUDO=True/i.test(oemR.out)
  const playbookName = oemR.out.match(/PLAYBOOKS=(.*)/)?.[1]?.trim() || ''

  // Combine all fingerprint strings for pattern matching
  const fingerprint = [oemMan, oemMod, oemUrl, regOrg, regOwn, playbookName].join(' ').toLowerCase()

  let osType = 'Stock'
  let osLabel = 'Standard Windows installation'

  const customOSPatterns = [
    // AME Wizard playbooks — matched by OEM fields + AME dir
    { test: () => /fsos/i.test(fingerprint),
      name: 'FSOS-X', label: `FSOS-X Playbook detected (${oemMod || regOrg})`, severity: 'high' },
    { test: () => /atlas/i.test(fingerprint) || hasAtlasMod || hasAtlasDir,
      name: 'AtlasOS', label: 'AtlasOS detected — heavily modified Windows', severity: 'high' },
    { test: () => /revi/i.test(fingerprint) || hasReviDir,
      name: 'ReviOS', label: 'ReviOS detected — custom lightweight Windows', severity: 'high' },
    { test: () => /tiny11|tiny10/i.test(fingerprint),
      name: 'Tiny11', label: 'Tiny11 detected — stripped Windows 11/10', severity: 'medium' },
    { test: () => /xlite/i.test(fingerprint),
      name: 'xLite', label: 'xLite detected — custom stripped Windows', severity: 'medium' },
    { test: () => /ghost.spectre|ghostspectre/i.test(fingerprint),
      name: 'GhostSpectre', label: 'GhostSpectre Superlite detected', severity: 'medium' },
    { test: () => /ameliorated|amelio/i.test(fingerprint),
      name: 'Windows Ameliorated', label: 'Windows Ameliorated Edition detected', severity: 'high' },
    { test: () => /framesynclabs|framesync/i.test(fingerprint),
      name: 'FrameSync Labs Playbook', label: `FrameSync Labs playbook detected (${oemMod || regOrg})`, severity: 'high' },
    // Generic AME wizard fallback — AME dir exists but no specific name matched
    { test: () => hasAMEConf && osType === 'Stock',
      name: 'AME Wizard Playbook', label: `AME Wizard playbook applied${playbookName ? ': ' + playbookName : ''}`, severity: 'high' },
    { test: () => /voided|vos\b/i.test(fingerprint),
      name: 'Void OS', label: 'VoidOS / Voided Windows detected', severity: 'medium' },
    { test: () => /rectified|rectify/i.test(fingerprint),
      name: 'RectifyOS', label: 'RectifyOS detected', severity: 'medium' },
  ]

  for (const p of customOSPatterns) {
    if (p.test()) {
      osType = 'Custom'
      osLabel = p.label
      detections.push({ name: p.name, detail: p.label, severity: p.severity })
      score -= p.severity === 'high' ? 40 : 20
      break // first match wins for osType/osLabel
    }
  }

  // NSudo presence is an additional signal (used by AME and other tools)
  if (hasNSudo && osType === 'Stock') {
    osType = 'Modified'
    osLabel = 'NSudo.exe found — system likely modified by a playbook or script'
    detections.push({ name: 'NSudo Present', detail: 'NSudo.exe detected in System32 — indicator of playbook or privilege-escalation tool usage.', severity: 'medium' })
    score -= 10
  }

  // Check for absent services that custom OSes remove
  const svcR = await runPS(`
    $missing = @()
    foreach ($s in @('SysMain','wuauserv','DiagTrack','WSearch','BITS','DPS')) {
      if (-not (Get-Service $s -EA SilentlyContinue)) { $missing += $s }
    }
    Write-Output ($missing -join ',')
  `)
  const missingSvcs = svcR.out.trim().split(',').filter(Boolean)
  if (missingSvcs.length >= 2) {
    if (osType === 'Stock') { osType = 'Modified'; osLabel = 'Multiple system services removed — likely a stripped Windows build' }
    detections.push({ name: 'Removed System Services', detail: `Missing: ${missingSvcs.join(', ')} — these are normally present on a stock install`, severity: 'medium' })
    score -= 15
  }

  // ── 2. Script signatures ─────────────────────────────────────────────────
  const envR = await runPS(`
    Write-Output "TITUS=$([System.Environment]::GetEnvironmentVariable('TITUS_TOOLS','Machine'))"
    Write-Output "SOPHIA=$([System.Environment]::GetEnvironmentVariable('SOPHIA_SCRIPT','Machine'))"
    Write-Output "CHRIS_PROFILE=$(Test-Path 'C:\\Chris Titus Tech\\winutil' -EA SilentlyContinue)"
    Write-Output "SOPHIAPKG=$(Test-Path '$env:ProgramFiles\\Sophia Script' -EA SilentlyContinue)"
    Write-Output "WINUTIL=$(Test-Path 'C:\\ProgramData\\winutil' -EA SilentlyContinue)"
  `)
  if (/TITUS=.+/.test(envR.out) || /CHRIS_PROFILE=True/i.test(envR.out) || /WINUTIL=True/i.test(envR.out)) {
    detections.push({ name: 'Chris Titus WinUtil Detected', detail: 'Chris Titus Tech winutil script signatures found. Some tweaks may be redundant or conflict.', severity: 'medium' })
    if (osType === 'Stock') osType = 'Modified'
    score -= 15
  }
  if (/SOPHIA=.+/.test(envR.out) || /SOPHIAPKG=True/i.test(envR.out)) {
    detections.push({ name: 'Sophia Script Detected', detail: 'Sophia Script installation detected. Windows settings may already be customized.', severity: 'medium' })
    if (osType === 'Stock') osType = 'Modified'
    score -= 15
  }

  // ── 3. BCD tweak signatures ──────────────────────────────────────────────
  const bcdR = await runPS(`
    try {
      $bcd = bcdedit /enum | Out-String
      Write-Output "DYNAMICTICK=$(if ($bcd -match 'disabledynamictick.*Yes') { 'yes' } else { 'no' })"
      Write-Output "PLATFORMTICK=$(if ($bcd -match 'useplatformtick.*Yes') { 'yes' } else { 'no' })"
      Write-Output "TSCSYNC=$(if ($bcd -match 'tscsyncpolicy.*enhanced') { 'yes' } else { 'no' })"
    } catch { Write-Output 'BCD_FAIL' }
  `, 10000)
  if (/DYNAMICTICK=yes/i.test(bcdR.out)) {
    detections.push({ name: 'Dynamic Tick Disabled (BCD)', detail: 'bcdedit disabledynamictick=yes detected. Timer tweak already applied.', severity: 'low' })
    score -= 5
  }
  if (/TSCSYNC=yes/i.test(bcdR.out)) {
    detections.push({ name: 'TSC Sync Enhanced (BCD)', detail: 'bcdedit tscsyncpolicy=enhanced already set. Timer tweak already applied.', severity: 'low' })
    score -= 3
  }

  // ── 4. Power plan ────────────────────────────────────────────────────────
  const ppR = await runPS(`
    $plans = powercfg /list
    if ($plans -match 'e9a42b02-d5df-448d-aa00-03f14749eb61') { Write-Output 'ULTIMATE' }
    elseif ($plans -match '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c') { Write-Output 'HIGH' }
    else { Write-Output 'OTHER' }
  `)
  const ppType = ppR.out.trim()
  if (ppType === 'ULTIMATE') {
    detections.push({ name: 'Ultimate Performance Plan', detail: 'Ultimate Performance power plan already active. Power tweak redundant.', severity: 'low' })
    score -= 3
  }

  // ── 5. Per-tweak state scan ──────────────────────────────────────────────
  const tweakScanR = await runPS(`
    function gp($p,$n){ try{(Get-ItemProperty $p -EA Stop).$n}catch{$null} }
    function svc($n){ $s=Get-Service $n -EA SilentlyContinue; if($s){[string]$s.StartType}else{'MISSING'} }
    function task($path,$name){ $t=Get-ScheduledTask -TaskPath $path -TaskName $name -EA SilentlyContinue; if($t){[string]$t.State}else{'MISSING'} }
    $r = @{}
    $gpu0 = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4d36e968-e325-11ce-bfc1-08002be10318}\\0000'
    $mm   = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management'
    $tcpP = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip\\Parameters'
    $bcd  = try { bcdedit /enum | Out-String } catch { '' }
    $pp   = powercfg /list 2>$null

    # ── Windows ─────────────────────────────────────────────────────────────
    $r['game-dvr']               = (gp 'HKCU:\\System\\GameConfigStore' 'GameDVR_Enabled') -eq 0
    $r['disable-fso']            = (gp 'HKCU:\\System\\GameConfigStore' 'GameDVR_FSEBehaviorMode') -eq 2
    $r['telemetry']              = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection' 'AllowTelemetry') -ne $null -and (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection' 'AllowTelemetry') -le 1
    $r['telemetry-zero']         = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection' 'AllowTelemetry') -eq 0
    $r['cortana']                = (svc 'WSearch') -eq 'Disabled'
    $r['win-tips']               = (gp 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager' 'SoftLandingEnabled') -eq 0
    $r['visual-effects']         = (gp 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects' 'VisualFXSetting') -eq 2
    $r['hibernate']              = -not (Test-Path 'C:\\hiberfil.sys')
    $r['disable-mouse-accel']    = (gp 'HKCU:\\Control Panel\\Mouse' 'MouseSpeed') -eq '0'
    $r['mmcss']                  = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games' 'GPU Priority') -ge 8
    $r['sysmain']                = (svc 'SysMain') -eq 'Disabled'
    $r['power-throttling']       = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling' 'PowerThrottlingOff') -eq 1
    $r['hpet']                   = $bcd -match 'useplatformclock.*No'
    $r['tsc-sync']               = $bcd -match 'tscsyncpolicy.*enhanced'
    $r['bcd-tweaks']             = $bcd -match 'bootmenupolicy.*Standard'
    $r['gpu-hwsch']              = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers' 'HwSchMode') -eq 2
    $r['ntfs']                   = (& fsutil behavior query disable8dot3 C: 2>$null) -match 'disabled'
    $mc = Get-MMAgent -EA SilentlyContinue
    $r['memory-compression']     = $mc -and -not $mc.MemoryCompression
    $r['tcp-stack']              = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched' 'NonBestEffortLimit') -eq 0
    $r['tcp-buffers']            = (gp $tcpP 'TcpAckFrequency') -eq 1
    $r['usb-suspend']            = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\USB' 'DisableSelectiveSuspend') -eq 1
    $r['disable-background-apps']    = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\AppPrivacy' 'LetAppsRunInBackground') -eq 2
    $r['disable-auto-maintenance']   = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Schedule\\Maintenance' 'MaintenanceDisabled') -eq 1
    $r['disable-delivery-opt']   = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization' 'DODownloadMode') -eq 0 -or (svc 'DoSvc') -eq 'Disabled'
    $r['priority-sep']           = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl' 'Win32PrioritySeparation') -eq 38
    $r['ultimate-perf']          = $pp -match 'e9a42b02-d5df-448d-aa00-03f14749eb61'
    $r['cpu-parking']            = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerSettings\\54533251-82be-4824-96c1-47b60b740d00\\0cc5b647-c1df-4637-891a-dec35c318583' 'ValueMax') -eq 0
    $r['spectre-meltdown']       = (gp $mm 'FeatureSettingsOverride') -eq 3
    $r['tdr-delay']              = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers' 'TdrDelay') -ge 10
    $r['search-cloud-off']       = (gp 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search' 'BingSearchEnabled') -eq 0
    $r['search-box-local']       = (gp 'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search' 'AllowSearchToUseLocation') -eq 0
    $r['cloud-sync-off']         = (svc 'OneDrive') -eq 'Disabled' -or -not (Test-Path "$env:LOCALAPPDATA\\Microsoft\\OneDrive\\OneDrive.exe")
    $r['bluetooth-hard-off']     = (svc 'bthserv') -eq 'Disabled'
    $r['boot-sound-off']         = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Authentication\\LogonUI\\BootAnimation' 'DisableStartupSound') -eq 1
    $r['cortana-blackout']       = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search' 'AllowCortana') -eq 0
    $r['defender-core-off']      = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender' 'DisableAntiSpyware') -eq 1
    $r['sticky-keys-guard']      = (gp 'HKCU:\\Control Panel\\Accessibility\\StickyKeys' 'Flags') -eq '506'
    $r['gpu-hags-off']           = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers' 'HwSchMode') -eq 1
    $r['sensor-suite-off']       = (svc 'SensorService') -eq 'Disabled'
    $r['diagnostics-off']        = (svc 'DPS') -eq 'Disabled'
    $r['biometrics-off']         = (svc 'WbioSrvc') -eq 'Disabled'
    $r['do-solo-mode']           = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\DeliveryOptimization' 'DODownloadMode') -eq 1
    $r['usb-power-guard']        = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\USB' 'DisableSelectiveSuspend') -eq 1
    $r['compat-scanner-off']     = (svc 'PcaSvc') -eq 'Disabled'
    $r['reality-telephony-off']  = (svc 'TapiSrv') -eq 'Disabled'
    $r['security-quiet']         = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows Defender Security Center\\Notifications' 'DisableNotifications') -eq 1
    $r['rt-scan-trim']           = $false
    $r['explorer-perf']          = (gp 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced' 'LaunchTo') -eq 1
    $r['win-update-defer']       = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU' 'NoAutoUpdate') -eq 1
    $r['nudge-blocker']          = (gp 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\ContentDeliveryManager' 'ContentDeliveryAllowed') -eq 0
    $r['cortana-remnant-off']    = -not (Get-AppxPackage -AllUsers -Name 'Microsoft.549981C3F5F10' -EA SilentlyContinue)
    $r['security-toast-off']     = (gp 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\Windows.Defender.SecurityCenter' 'Enabled') -eq 0
    $r['proximity-deny']         = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\proximity' 'Value') -eq 'Deny'
    $r['capability-ruleset']     = (gp 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Privacy' 'TailoredExperiencesWithDiagnosticDataEnabled') -eq 0
    $r['diag-tasks-off']         = (task '\\Microsoft\\Windows\\Diagnosis\\' 'Scheduled') -eq 'Disabled'
    $r['compat-tasks-off']       = (task '\\Microsoft\\Windows\\Application Experience\\' 'StartupAppTask') -eq 'Disabled'
    $r['cursor-max-rate']        = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\mouclass\\Parameters' 'MouseDataQueueSize') -eq 20
    $r['cpu-focus-bias']         = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl' 'Win32PrioritySeparation') -eq 38
    $r['net-throttling']         = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile' 'NetworkThrottlingIndex') -eq 0xFFFFFFFF
    $r['disable-netbios']        = $false  # per-adapter check — complex, skip
    $r['disable-ipv6']           = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\Tcpip6\\Parameters' 'DisabledComponents') -eq 255
    $r['disable-wpad']           = (svc 'WinHttpAutoProxySvc') -eq 'Disabled'
    $r['qos-reserve']            = (gp 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched' 'NonBestEffortLimit') -eq 0
    $r['dns-cloudflare']         = (Get-DnsClientServerAddress -EA SilentlyContinue | Where-Object {$_.ServerAddresses -contains '1.1.1.1'} | Select-Object -First 1) -ne $null
    $r['dns-google']             = (Get-DnsClientServerAddress -EA SilentlyContinue | Where-Object {$_.ServerAddresses -contains '8.8.8.8'} | Select-Object -First 1) -ne $null

    # ── Hardware ─────────────────────────────────────────────────────────────
    $r['nvidia-max-perf']        = (gp 'HKLM:\\SOFTWARE\\NVIDIA Corporation\\Global\\NvTweak' 'Powersave') -eq 0
    $r['nvidia-low-latency']     = (gp $gpu0 'D3PCLatency') -eq 1
    $r['disable-hdcp']           = (gp $gpu0 'RMHdcpKeyglobZero') -eq 1
    $r['nvidia-pstate-lock']     = (gp $gpu0 'DisableDynamicPstates') -eq 1
    $r['nvidia-profile']         = (gp $gpu0 'RMAERQSizeMultiplier') -eq 32
    $r['amd-chill-off']          = (gp $gpu0 'KMD_EnableChill') -eq 0
    $r['amd-power-hold']         = (gp $gpu0 'KMD_EnableABC') -eq 0
    $r['amd-service-trim']       = (svc 'amdfendr') -eq 'Disabled'
    $r['corsair-audio-off']      = (svc 'CorsairGamingAudioConfig') -eq 'Disabled'
    $r['hvci-offload']           = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceGuard\\Scenarios\\HypervisorEnforcedCodeIntegrity' 'Enabled') -eq 0
    $r['memory-guard-tune']      = (gp $mm 'DisablePagingExecutive') -eq 1
    $r['full-mitigation-wipe']   = (gp $mm 'FeatureSettingsOverride') -eq 3 -and (gp $mm 'FeatureSettingsOverrideMask') -eq 3
    $r['process-count-reduction'] = do {
      $ramKB = [math]::Round((Get-WmiObject Win32_ComputerSystem -EA SilentlyContinue).TotalPhysicalMemory / 1024)
      $val = (gp $mm 'SvcHostSplitThresholdInKB')
      $val -ne $null -and $val -ge ($ramKB * 0.9)
    }
    $r['nvme-latency']           = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\StorPort' 'TotalRequestHoldOffPeriod') -eq 0
    $r['speed-shift']            = (gp 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Power' 'EnergyEstimationEnabled') -eq 0
    $r['apex-power-plan']        = $pp -match 'Apex'

    # MSI mode — sample first 5 non-network PCI devices for MSISupported=1
    $r['msi-mode'] = do {
      $netGuid = '{4d36e972-e325-11ce-bfc1-08002be10318}'
      $found = $false
      $checked = 0
      Get-ChildItem 'HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\PCI' -EA SilentlyContinue | ForEach-Object {
        if ($checked -ge 5) { return }
        Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
          if ($checked -ge 5) { return }
          $props = Get-ItemProperty -Path $_.PSPath -EA SilentlyContinue
          if ($props.ClassGUID -eq $netGuid) { return }
          $msiPath = "$($_.PSPath)\\Device Parameters\\Interrupt Management\\MessageSignaledInterruptProperties"
          if ((gp $msiPath 'MSISupported') -eq 1) { $found = $true }
          $checked++
        }
      }
      $found
    }

    # NIC offloads — check first up Ethernet adapter
    $r['nic-offloads'] = do {
      $nic = Get-NetAdapter | Where-Object {
        $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
        $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*'
      } | Select-Object -First 1
      if ($nic) {
        $prop = Get-NetAdapterAdvancedProperty -Name $nic.Name -RegistryKeyword '*IPChecksumOffloadIPv4*' -EA SilentlyContinue
        $prop -and $prop.DisplayValue -eq 'Disabled'
      } else { $false }
    }

    $r['nic-interrupt-mod'] = do {
      $nic = Get-NetAdapter | Where-Object {
        $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
        $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*'
      } | Select-Object -First 1
      if ($nic) {
        $prop = Get-NetAdapterAdvancedProperty -Name $nic.Name -RegistryKeyword '*InterruptModeration*' -EA SilentlyContinue
        $prop -and $prop.DisplayValue -eq 'Disabled'
      } else { $false }
    }

    $r['nic-flow-control'] = do {
      $nic = Get-NetAdapter | Where-Object {
        $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
        $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*'
      } | Select-Object -First 1
      if ($nic) {
        $prop = Get-NetAdapterAdvancedProperty -Name $nic.Name -RegistryKeyword '*FlowControl*' -EA SilentlyContinue
        $prop -and $prop.DisplayValue -eq 'Disabled'
      } else { $false }
    }

    $r['nic-energy-efficient'] = do {
      $nic = Get-NetAdapter | Where-Object {
        $_.Status -eq 'Up' -and $_.PhysicalMediaType -ne 'Native 802.11' -and
        $_.InterfaceDescription -notlike '*Wireless*' -and $_.InterfaceDescription -notlike '*Wi-Fi*'
      } | Select-Object -First 1
      if ($nic) {
        $prop = Get-NetAdapterAdvancedProperty -Name $nic.Name -RegistryKeyword '*EEE*' -EA SilentlyContinue
        $prop -and $prop.DisplayValue -eq 'Disabled'
      } else { $false }
    }

    # Defender quiet
    $mp = Get-MpPreference -EA SilentlyContinue
    $r['defender-quiet']         = $mp -and $mp.MAPSReporting -eq 0

    # ── FiveM ────────────────────────────────────────────────────────────────
    $fivemIni = "$env:LOCALAPPDATA\\FiveM\\FiveM.app\\CitizenFX.ini"
    $iniRaw   = if (Test-Path $fivemIni) { Get-Content $fivemIni -Raw -EA SilentlyContinue } else { '' }
    $ifeo     = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
    $r['fivem-priority']         = (gp "$ifeo\\FiveM.exe\\PerfOptions" 'CpuPriorityClass') -eq 3
    $r['fivem-io-priority']      = (gp "$ifeo\\FiveM.exe\\PerfOptions" 'IoPriority') -eq 3
    $r['fivem-defender']         = $mp -and ($mp.ExclusionPath -contains "$env:LOCALAPPDATA\\FiveM")
    $r['fivem-hang-fix']         = (gp 'HKCU:\\Control Panel\\Desktop' 'HungAppTimeout') -eq '3000'
    $r['fivem-network']          = (gp $tcpP 'MaxUserPort') -eq 65534
    $r['fivem-mmcss']            = (gp 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games' 'GPU Priority') -ge 8
    $r['fivem-fso']              = (gp 'HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers' 'FiveM.exe') -like '*DISABLEDXMAXIMIZEDWINDOWEDMODE*'
    $r['fivem-gpu']              = (gp $gpu0 'F1MaxLatency') -eq 0
    $r['fivem-vm']               = (gp $mm 'DisablePagingExecutive') -eq 1
    $r['fivem-streaming-mem']    = $iniRaw -match 'StreamingMemory\s*=\s*1024'
    $r['fivem-worker-threads']   = $iniRaw -match 'WorkerThreads\s*=\s*\d+'
    $r['fivem-disable-crash-reporter']      = $iniRaw -match 'DisableCrashReporter\s*=\s*1'
    $r['fivem-disable-anticheat-upload']    = $iniRaw -match 'DisableSteamAchievements\s*=\s*true'
    $r['fivem-disable-update-checks']       = $iniRaw -match 'UpdateChannel\s*=\s*canary'
    $r['fivem-preload-ipl']      = $iniRaw -match 'MaximumGrass\s*=\s*0'
    $r['fivem-reduce-draw-distance'] = $iniRaw -match 'MaxLodDistance\s*=\s*30'

    # Emit KEY=VALUE lines
    foreach ($k in $r.Keys) {
      Write-Output "$k=$(if($r[$k]){'1'}else{'0'})"
    }
  `, 45000)

  // Parse into { id -> bool }
  const tweakStates = {}
  for (const line of tweakScanR.out.split(/\r?\n/)) {
    const m = line.match(/^([a-z0-9-]+)=([01])$/)
    if (m) tweakStates[m[1]] = m[2] === '1'
  }

  // Build legacy conflicts list for existing UI (tweaks that are already applied)
  const TWEAK_LABELS = {
    'game-dvr': 'Game DVR Disabled', 'disable-fso': 'FSO Disabled', 'telemetry': 'Telemetry Off',
    'cortana': 'WSearch Disabled', 'visual-effects': 'Visual Effects Minimal', 'hibernate': 'Hibernate Off',
    'disable-mouse-accel': 'Mouse Accel Off', 'mmcss': 'MMCSS Optimized', 'sysmain': 'SysMain Disabled',
    'power-throttling': 'Power Throttling Off', 'hpet': 'HPET Disabled', 'tsc-sync': 'TSC Sync Enhanced',
    'gpu-hwsch': 'GPU HW Scheduling On', 'ntfs': 'NTFS 8.3 Off', 'memory-compression': 'Memory Compression Off',
    'tcp-buffers': 'Nagle Off', 'ultimate-perf': 'Ultimate Performance Plan', 'spectre-meltdown': 'Mitigations Disabled',
    'bluetooth-hard-off': 'Bluetooth Disabled', 'defender-core-off': 'Defender Core Off',
    'cortana-blackout': 'Cortana Blackout', 'search-cloud-off': 'Bing Search Off',
  }
  const conflicts = Object.entries(tweakStates)
    .filter(([id, applied]) => applied && TWEAK_LABELS[id])
    .map(([id]) => TWEAK_LABELS[id])

  clearDiscordOverride()
  return {
    osType,
    osLabel,
    conflicts,
    detections,
    tweakStates,
    safetyScore: Math.max(0, Math.min(100, score))
  }
})

// ─── Auto-Updater ─────────────────────────────────────────────────────────────
let autoUpdater = null
try {
  const { autoUpdater: au } = require('electron-updater')
  autoUpdater = au
  autoUpdater.autoDownload = false
  autoUpdater.autoInstallOnAppQuit = true
  autoUpdater.logger = require('electron-log')
  autoUpdater.logger.transports.file.level = 'info'

  autoUpdater.on('update-available', (info) => {
    mainWindow?.webContents.send('update-available', info)
  })
  autoUpdater.on('update-not-available', () => {
    mainWindow?.webContents.send('update-not-available')
  })
  autoUpdater.on('download-progress', (progress) => {
    mainWindow?.webContents.send('update-progress', progress)
  })
  autoUpdater.on('update-downloaded', (info) => {
    mainWindow?.webContents.send('update-downloaded', info)
  })
  autoUpdater.on('error', (err) => {
    mainWindow?.webContents.send('update-error', err?.message)
  })
} catch (e) {
  // electron-updater not installed yet — silently skip
}

ipcMain.handle('check-update', async () => {
  if (!autoUpdater) return { ok: false, reason: 'electron-updater not installed' }
  try { await autoUpdater.checkForUpdates(); return { ok: true } }
  catch (e) { return { ok: false, reason: e.message } }
})

ipcMain.handle('download-update', async () => {
  if (!autoUpdater) return { ok: false }
  try { await autoUpdater.downloadUpdate(); return { ok: true } }
  catch (e) { return { ok: false, reason: e.message } }
})

ipcMain.handle('install-update', () => {
  autoUpdater?.quitAndInstall(false, true)
})

// What's New content
const WHATS_NEW = [
  { version: '1.1', date: 'May 2026', items: [
    'System tray — app minimizes to tray instead of taskbar; close button hides to tray; single-click to restore',
    'Tray context menu — shows Pulse status at a glance and lets you toggle Pulse on/off without opening the app',
    'Auto-Pulse — Game Watcher can now automatically activate Pulse when a game is detected and restore when it closes',
    'Auto-Pulse badge — a green "AUTO-PULSE ACTIVE" indicator appears on the Game Watcher card while active, with a manual stop button',
    'Pulse toggle from tray — start or stop Pulse directly from the right-click tray menu; last-used preset is remembered',
    'Pre-Flight scan added to first-launch wizard — detects pre-applied tweaks before Auto-Optimize runs',
    'Startup & Services tab — renamed from Services, added a Recommended button that auto-disables known bloatware startup items',
    'Cleanup tab expanded — 13 cleanup categories including browser caches, app caches, WinSxS, event logs, thumbnail cache, FiveM cache, and more',
    'FiveM settings save reliability — warning banner now shown before first save; instructions to avoid re-opening Graphics Settings in-game',
    'Fixed: Discord Rich Presence not showing when Discord is already open at launch',
    'Fixed: Discord status lingering after app close — clearActivity now awaited before socket teardown',
    'Fixed: BCDEdit tweak card and Arma Reforger Pulse preset showing no icon',
  ]},
  { version: '1.0', date: 'May 2026', items: [
    'Jylli Tool — full rebrand with a cleaner, faster interface',
    'Fixes page — one-click fixes for common Windows issues (WiFi, Bluetooth, Valorant CFG, NVIDIA Control Panel, etc.)',
    'Pre-Flight Diagnostic — detects custom OS (AtlasOS, ReviOS, Tiny11), Chris Titus & Sophia scripts, and existing tweaks before applying',
    'New Advanced tweaks: GPU HAGS Off, HVCI Offload, Nvidia P-State Lock, Memory Guard, Full Mitigation Wipe, Apex Power Plan',
    'New General tweaks: Raw Aim Curve, Proximity Deny, Reality/Telephony Off, Cloud Sync Off, DO Solo Mode, Nudge Blocker, and more',
    'FiveM in-game settings completely reworked — now reliably finds prefs.xml and offers manual path override',
    'Safety warnings added to all high-risk tweaks',
    'Pulse — dedicated real-time game optimizer window with live CPU, RAM & GPU monitoring',
    'Game Profiles — detects your installed games and applies per-game tweaks',
    'App Optimizer — cut background resource usage from browsers, Discord, Spotify and more',
  ]},
]

ipcMain.handle('get-whats-new', () => WHATS_NEW)

// ─── App Optimizer IPC ────────────────────────────────────────────────────────
const APP_OPTIMIZER_TWEAKS = {
  // ── Browsers ─────────────────────────────────────────────────────────────
  'edge-optimize': {
    apply: async (s, ps, cmd) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Edge' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Edge' -Name 'StartupBoostEnabled' -Value 0 -Force
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Edge' -Name 'BackgroundModeEnabled' -Value 0 -Force
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Edge' -Name 'HardwareAccelerationModeEnabled' -Value 0 -Force`)
      s('Edge: startup boost off, background mode off, HW accel disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Remove-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Edge' -Name 'StartupBoostEnabled','BackgroundModeEnabled' -EA SilentlyContinue`)
      s('Edge settings restored.', 'ok')
    }
  },
  'chrome-optimize': {
    apply: async (s, ps) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Policies\\Google\\Chrome' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Google\\Chrome' -Name 'HardwareAccelerationModeEnabled' -Value 0 -Force
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Google\\Chrome' -Name 'BackgroundModeEnabled' -Value 0 -Force`)
      s('Chrome: hardware acceleration off, background mode off.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Remove-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Google\\Chrome' -Name 'HardwareAccelerationModeEnabled','BackgroundModeEnabled' -EA SilentlyContinue`)
      s('Chrome settings restored.', 'ok')
    }
  },
  'firefox-optimize': {
    apply: async (s, ps) => {
      const ffProfiles = path.join(os.homedir(), 'AppData', 'Roaming', 'Mozilla', 'Firefox', 'Profiles')
      if (fs.existsSync(ffProfiles)) {
        const userJs = `// PTU Firefox Performance
user_pref("layers.acceleration.disabled", false);
user_pref("gfx.webrender.all", true);
user_pref("media.hardware-video-decoding.enabled", true);
user_pref("browser.cache.memory.capacity", 524288);
user_pref("network.prefetch-next", false);
user_pref("network.dns.disablePrefetch", true);
user_pref("browser.sessionstore.interval", 60000);
user_pref("browser.sessionhistory.max_total_viewers", 4);`
        const profiles = fs.readdirSync(ffProfiles)
        for (const p of profiles) {
          const profilePath = path.join(ffProfiles, p)
          if (fs.statSync(profilePath).isDirectory()) {
            fs.writeFileSync(path.join(profilePath, 'user.js'), userJs)
            s(`  Firefox profile ${p}: user.js written`, 'ok')
          }
        }
      } else { s('Firefox not found.', 'info') }
    },
    restore: async (s) => s('Delete user.js from Firefox profile folder to restore.', 'info')
  },
  'brave-optimize': {
    apply: async (s, ps) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Policies\\BraveSoftware\\Brave' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\BraveSoftware\\Brave' -Name 'HardwareAccelerationModeEnabled' -Value 0 -Force
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\BraveSoftware\\Brave' -Name 'BackgroundModeEnabled' -Value 0 -Force`)
      s('Brave: hardware acceleration and background mode disabled.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Remove-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\BraveSoftware\\Brave' -Name 'HardwareAccelerationModeEnabled','BackgroundModeEnabled' -EA SilentlyContinue`)
      s('Brave restored.', 'ok')
    }
  },
  'opera-gx-optimize': {
    apply: async (s, ps) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Policies\\OperaSoftware\\OperaGX' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\OperaSoftware\\OperaGX' -Name 'HardwareAccelerationModeEnabled' -Value 0 -Force`)
      s('Opera GX: hardware acceleration disabled.', 'ok')
    },
    restore: async (s, ps) => { s('Opera GX: no changes to restore.', 'info') }
  },
  // ── Discord ───────────────────────────────────────────────────────────────
  'discord-optimize': {
    apply: async (s, ps) => {
      const appData = process.env.APPDATA || ''
      const discordCfg = path.join(appData, 'discord', 'settings.json')
      if (fs.existsSync(discordCfg)) {
        const cfg = JSON.parse(fs.readFileSync(discordCfg, 'utf8'))
        cfg.SKIP_HOST_UPDATE = true
        cfg.enableHardwareAcceleration = false
        cfg.audioSubsystem = 'legacy'
        fs.writeFileSync(discordCfg, JSON.stringify(cfg, null, 2))
        s('Discord: HW accel off, skip host update, legacy audio subsystem.', 'ok')
      } else { s('Discord settings.json not found — launch Discord first.', 'warn') }
      // Reduce Discord process priority
      await ps(`Get-Process -Name 'Discord','DiscordPTB','DiscordCanary' -EA SilentlyContinue | ForEach-Object { $_.PriorityClass = 'BelowNormal' }`)
    },
    restore: async (s) => s('Restart Discord to reset to defaults.', 'info')
  },
  'discord-cpu-priority': {
    apply: async (s, ps) => {
      const base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
      await ps(`New-Item -Path "${base}\\Discord.exe\\PerfOptions" -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path "${base}\\Discord.exe\\PerfOptions" -Name CpuPriorityClass -Value 1 -Force
Set-ItemProperty -Path "${base}\\Discord.exe\\PerfOptions" -Name IoPriority -Value 0 -Force`)
      s('Discord.exe: IDLE CPU + I/O priority (lowest).', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\Discord.exe\\PerfOptions' -Name CpuPriorityClass -Value 2 -Force -EA SilentlyContinue`)
      s('Discord priority restored to Normal.', 'ok')
    }
  },
  // ── Spotify ───────────────────────────────────────────────────────────────
  'spotify-optimize': {
    apply: async (s, ps) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\Spotify.exe\\PerfOptions' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\Spotify.exe\\PerfOptions' -Name CpuPriorityClass -Value 1 -Force`)
      s('Spotify: IDLE CPU priority — won\'t compete with games.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\Spotify.exe\\PerfOptions' -Name CpuPriorityClass -Value 2 -Force -EA SilentlyContinue`)
      s('Spotify priority restored.', 'ok')
    }
  },
  // ── Steam ─────────────────────────────────────────────────────────────────
  'steam-optimize': {
    apply: async (s, ps, cmd) => {
      const steamPath = await ps('(Get-ItemProperty "HKLM:\\SOFTWARE\\WOW6432Node\\Valve\\Steam" -EA SilentlyContinue).InstallPath')
      if (steamPath.out.trim()) {
        const cfgPath = path.join(steamPath.out.trim(), 'steam.cfg')
        fs.writeFileSync(cfgPath, 'BootStrapperInhibitAll=enable\nCleintAutoRestart=disable\n')
        s('Steam: bootstrap inhibited, auto-restart disabled.', 'ok')
      }
      await ps(`Set-ItemProperty -Path 'HKCU:\\Software\\Valve\\Steam' -Name 'StartupMode' -Value 0 -Force -EA SilentlyContinue`)
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\steam.exe\\PerfOptions' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\steam.exe\\PerfOptions' -Name CpuPriorityClass -Value 1 -Force`)
      s('Steam: startup minimized, web helper priority reduced.', 'ok')
    },
    restore: async (s) => s('Delete steam.cfg from Steam install folder to restore.', 'info')
  },
  // ── OBS ───────────────────────────────────────────────────────────────────
  'obs-optimize': {
    apply: async (s, ps) => {
      await ps(`New-Item -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\obs64.exe\\PerfOptions' -Force -EA SilentlyContinue | Out-Null
Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\obs64.exe\\PerfOptions' -Name CpuPriorityClass -Value 3 -Force`)
      s('OBS: HIGH CPU priority — encoding gets more CPU time.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\obs64.exe\\PerfOptions' -Name CpuPriorityClass -Value 2 -Force -EA SilentlyContinue`)
      s('OBS priority restored.', 'ok')
    }
  },
  // ── Windows Explorer ──────────────────────────────────────────────────────
  'explorer-optimize': {
    apply: async (s, ps) => {
      await ps(`Set-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced' -Name 'LaunchTo' -Value 1 -Force
Set-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced' -Name 'HideFileExt' -Value 0 -Force
Set-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize' -Name 'StartupDelayInMSec' -Value 0 -Force -EA SilentlyContinue`)
      s('Explorer: no startup delay, launches to This PC, show extensions.', 'ok')
    },
    restore: async (s, ps) => {
      await ps(`Remove-ItemProperty -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize' -Name 'StartupDelayInMSec' -EA SilentlyContinue`)
      s('Explorer tweaks restored.', 'ok')
    }
  },
}

ipcMain.handle('run-app-optimizer', async (_, { id, action }) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  send(`App Optimizer: ${id} [${action}]`, 'head')
  try {
    await APP_OPTIMIZER_TWEAKS[id]?.[action]?.(send, runPS, runCmd)
    return { ok: true }
  } catch (e) {
    send(`Error: ${e.message}`, 'err')
    return { ok: false, error: e.message }
  }
})
const GAME_EXE_MAP = {
  'arma-reforger': ['ArmaReforger.exe', 'ArmaReforgerSteam.exe', 'ArmaReforger_BE.exe'],
  'fortnite':      ['FortniteClient-Win64-Shipping.exe', 'FortniteLauncher.exe', 'EpicGamesLauncher.exe', 'FortniteClient-Win64-Shipping_BE.exe'],
  'gtav':          ['GTA5.exe', 'GTAVLauncher.exe', 'PlayGTAV.exe', 'FiveM.exe'],
  'valorant':      ['VALORANT-Win64-Shipping.exe', 'RiotClientServices.exe', 'vgc.exe'],
  'minecraft':     ['javaw.exe', 'java.exe', 'Minecraft.Windows.exe', 'MinecraftLauncher.exe'],
  'cs2':           ['cs2.exe', 'csgo.exe'],
  'apex':          ['r5apex.exe', 'r5apex_dx12.exe', 'EADesktop.exe', 'Origin.exe'],
  'warzone':       ['cod.exe', 'ModernWarfare.exe', 'Warzone.exe', 'battle.net.exe'],
  'rust':          ['RustClient.exe', 'rust.exe'],
  'the-finals':    ['FINALS-Win64-Shipping.exe', 'Discovery.exe'],
}

ipcMain.handle('apply-game-tweak', async (_, { gameId, action }) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  const exes = GAME_EXE_MAP[gameId] || []
  const ifeoBase = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options'
  const mmcssBase = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'

  send(`◆ Game Tweak [${gameId}] — ${action}`, 'head')

  if (action === 'apply') {
    // Set HIGH CPU + IO priority for each exe via IFEO
    for (const exe of exes) {
      await runPS(`New-Item -Path "${ifeoBase}\\${exe}\\PerfOptions" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "${ifeoBase}\\${exe}\\PerfOptions" -Name CpuPriorityClass -Value 3 -Force; Set-ItemProperty -Path "${ifeoBase}\\${exe}\\PerfOptions" -Name IoPriority -Value 3 -Force`)
      send(`  ${exe}: HIGH CPU + IO priority`, 'ok')
    }
    // MMCSS GPU priority
    await runPS(`Set-ItemProperty -Path "${mmcssBase}" -Name SystemResponsiveness -Value 10 -Force`)
    await runPS(`Set-ItemProperty -Path "${mmcssBase}\\Tasks\\Games" -Name "GPU Priority" -Value 8 -Force`)
    await runPS(`Set-ItemProperty -Path "${mmcssBase}\\Tasks\\Games" -Name "Priority" -Value 6 -Force`)
    await runPS(`Set-ItemProperty -Path "${mmcssBase}\\Tasks\\Games" -Name "Scheduling Category" -Value "High" -Force`)
    await runPS(`Set-ItemProperty -Path "${mmcssBase}\\Tasks\\Games" -Name "Clock Rate" -Value 10000 -Force`)
    send('MMCSS game scheduling applied.', 'ok')

    // Disable fullscreen optimizations for each exe
    const compat = 'HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers'
    for (const exe of exes) {
      await runPS(`New-Item -Path "${compat}" -Force -EA SilentlyContinue | Out-Null; Set-ItemProperty -Path "${compat}" -Name "${exe}" -Value "~ DISABLEDXMAXIMIZEDWINDOWEDMODE" -Force`)
    }
    send('Fullscreen Optimizations disabled for game exes.', 'ok')

    // Game DVR off
    await runPS('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_Enabled -Value 0 -Force')
    send('Game DVR disabled.', 'ok')

    send(`✓ ${gameId} optimizations applied!`, 'ok')
  } else {
    // Restore: set Normal priority
    for (const exe of exes) {
      await runPS(`Set-ItemProperty -Path "${ifeoBase}\\${exe}\\PerfOptions" -Name CpuPriorityClass -Value 2 -Force -EA SilentlyContinue`)
      await runPS(`Set-ItemProperty -Path "${ifeoBase}\\${exe}\\PerfOptions" -Name IoPriority -Value 2 -Force -EA SilentlyContinue`)
    }
    send(`${gameId} priorities restored to Normal.`, 'ok')
  }

  return { ok: true }
})

// Clean temp files
ipcMain.handle('clean-system', async (_, paths) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  for (const p of paths) {
    const r = await runPS(`Remove-Item -Path '${p}\\*' -Recurse -Force -ErrorAction SilentlyContinue; (Get-ChildItem '${p}' -EA SilentlyContinue).Count`)
    send(`  ${path.basename(p)}: cleared`, 'ok')
  }
  send('Cleanup complete.', 'ok')
  return { ok: true }
})

// ─── Process Manager ──────────────────────────────────────────────────────────
ipcMain.handle('get-processes', async () => {
  const r = await runPS(`
    Get-Process | Where-Object {$_.CPU -ne $null} |
    Select-Object Name, Id,
      @{N='CPU';E={[math]::Round($_.CPU,1)}},
      @{N='RAM';E={[math]::Round($_.WorkingSet64/1MB,1)}},
      @{N='Path';E={try{$_.MainModule.FileName}catch{''}}} |
    Sort-Object RAM -Descending | Select-Object -First 80 |
    ConvertTo-Json -Depth 2
  `, 15000)
  try { return JSON.parse(r.out) } catch { return [] }
})

ipcMain.handle('kill-process', async (_, pid) => {
  const r = await runPS(`Stop-Process -Id ${pid} -Force -ErrorAction SilentlyContinue`)
  return { ok: r.ok }
})

ipcMain.handle('game-boost-apply', async () => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  send('◆ Game Boost — killing background processes…', 'head')

  const BOOST_KILL = [
    'Discord', 'DiscordPTB', 'DiscordCanary',
    'EpicGamesLauncher', 'EpicWebHelper',
    'steam', 'steamwebhelper', 'GameOverlayUI',
    'OneDrive', 'OneDriveStandaloneUpdater',
    'SearchApp', 'SearchHost',
    'Cortana', 'SearchUI',
    'Teams', 'ms-teams',
        'chrome', 'msedge', 'firefox', 'opera',
    'MicrosoftEdgeUpdate', 'GoogleUpdate',
    'SkypeApp', 'skype',
    'XboxApp', 'GamingServices', 'XblGameSave',
    'WinRAR', '7zG', 'notepad', 'notepad++',
    'msiAfterburner', 'RTSS',
  ]

  const killed = []
  for (const name of BOOST_KILL) {
    const r = await runPS(`
      $p = Get-Process -Name '${name}' -ErrorAction SilentlyContinue
      if ($p) { Stop-Process -Name '${name}' -Force -ErrorAction SilentlyContinue; Write-Output 'killed' }
    `)
    if (r.out.trim() === 'killed') {
      killed.push(name)
      send(`  ✓ Stopped: ${name}`, 'ok')
    }
  }

  // Set current process priorities for performance
  await runPS('Get-Process | Where-Object {$_.Name -notlike "System*" -and $_.Name -ne "Idle"} | ForEach-Object { try { $_.PriorityClass = "BelowNormal" } catch {} }')
  send(`Game Boost complete — stopped ${killed.length} background processes.`, 'ok')
  send('Click "End Boost" when done gaming to restore processes.', 'info')
  return { ok: true, killed }
})

ipcMain.handle('game-boost-restore', async () => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  send('Restoring background processes…', 'head')
  // Restore process priorities
  await runPS('Get-Process | Where-Object {$_.Name -notlike "System*"} | ForEach-Object { try { $_.PriorityClass = "Normal" } catch {} }')
  send('Process priorities restored. Restart apps manually if needed.', 'ok')
  return { ok: true }
})

// ─── Pulse Window ─────────────────────────────────────────────────────────────
ipcMain.handle('open-pulse-window', () => {
  if (pulseWindow && !pulseWindow.isDestroyed()) {
    pulseWindow.focus()
    return
  }
  pulseWindow = new BrowserWindow({
    width: 520,
    height: 780,
    minWidth: 480,
    minHeight: 600,
    frame: false,
    transparent: false,
    backgroundColor: '#0d0d0d',
    show: false,
    title: 'Pulse',
    icon: path.join(__dirname, 'assets', 'icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })
  pulseWindow.loadFile('pulse.html')
  pulseWindow.once('ready-to-show', () => pulseWindow.show())
  pulseWindow.on('closed', () => { pulseWindow = null })
})

ipcMain.handle('close-pulse-window', () => {
  if (pulseWindow && !pulseWindow.isDestroyed()) pulseWindow.close()
  pulseWindow = null
})

// ─── FiveM Graphics Settings Window ──────────────────────────────────────────
ipcMain.handle('open-fivem-settings', () => {
  if (fivemSettingsWindow && !fivemSettingsWindow.isDestroyed()) {
    fivemSettingsWindow.focus()
    return
  }
  fivemSettingsWindow = new BrowserWindow({
    width: 680,
    height: 740,
    minWidth: 560,
    minHeight: 500,
    frame: false,
    transparent: false,
    backgroundColor: '#0d0d0d',
    show: false,
    title: 'FiveM Graphics Settings',
    icon: path.join(__dirname, 'assets', 'icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })
  fivemSettingsWindow.loadFile('fivem-settings.html')
  fivemSettingsWindow.once('ready-to-show', () => fivemSettingsWindow.show())
  fivemSettingsWindow.on('closed', () => { fivemSettingsWindow = null })
})

ipcMain.handle('close-fivem-settings', () => {
  if (fivemSettingsWindow && !fivemSettingsWindow.isDestroyed()) fivemSettingsWindow.close()
  fivemSettingsWindow = null
})

// ─── Pulse ────────────────────────────────────────────────────────────────────
const PULSE_PRESETS = {
  fortnite:     { name: 'Fortnite',            exe: 'FortniteClient-Win64-Shipping', priority: 'High',        killList: ['EpicGamesLauncher','EpicWebHelper','chrome','msedge'] },
  valorant:     { name: 'Valorant',            exe: 'VALORANT-Win64-Shipping',       priority: 'High',        killList: ['chrome','msedge','OneDrive'] },
  csgo:         { name: 'CS2 / CS:GO',         exe: 'cs2',                           priority: 'High',        killList: ['chrome','msedge','steam','steamwebhelper'] },
  fivem:        { name: 'FiveM',               exe: 'FiveM_GTAProcess',              priority: 'AboveNormal', killList: ['chrome','msedge','OneDrive','Teams'] },
  warzone:      { name: 'Warzone',             exe: 'cod',                           priority: 'High',        killList: ['chrome','msedge','OneDrive'] },
  apex:         { name: 'Apex Legends',        exe: 'r5apex',                        priority: 'High',        killList: ['chrome','msedge','EpicGamesLauncher'] },
  minecraft:    { name: 'Minecraft',           exe: 'javaw',                         priority: 'AboveNormal', killList: ['chrome','msedge','OneDrive'] },
  arma:         { name: 'Arma Reforger',       exe: 'ArmaReforger*',                 priority: 'High',        killList: ['chrome','msedge','OneDrive','Teams'] },
  tarkov:       { name: 'Escape from Tarkov',  exe: 'EscapeFromTarkov',              priority: 'High',        killList: ['chrome','msedge','OneDrive'] },
  rocketleague: { name: 'Rocket League',       exe: 'RocketLeague',                  priority: 'High',        killList: ['chrome','msedge','EpicGamesLauncher'] },
  rust:         { name: 'Rust',                exe: 'RustClient',                    priority: 'High',        killList: ['chrome','msedge','OneDrive'] },
  r6:           { name: 'Rainbow Six Siege',   exe: 'RainbowSix',                    priority: 'High',        killList: ['chrome','msedge','OneDrive','Teams'] },
  overwatch:    { name: 'Overwatch 2',         exe: 'Overwatch',                     priority: 'High',        killList: ['chrome','msedge','OneDrive'] },
}

// Saved originals restored on pulse-stop
let pulseActivePreset = null
let pulseOrigMmcssResponsiveness = null
let pulseOrigPowerPlan = null
let autoPulseTriggeredPreset = null  // set when game watcher auto-activated pulse
let lastUsedPulsePreset = null       // remembered for tray "Start Pulse" shortcut


async function pulseRestoreAll(send) {
  // Restore timer resolution — just release our request; Windows reverts automatically
  await runPS(`
    try {
      Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class TimerRes { [DllImport("ntdll.dll")] public static extern int NtSetTimerResolution(uint r, bool set, out uint cur); }
'@
      $cur = [uint32]0
      [TimerRes]::NtSetTimerResolution(156250, $false, [ref]$cur) | Out-Null
    } catch {}
  `, 8000)

  // Restore MMCSS SystemResponsiveness
  if (pulseOrigMmcssResponsiveness !== null) {
    await runPS(`Set-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile' -Name SystemResponsiveness -Value ${pulseOrigMmcssResponsiveness} -Force -EA SilentlyContinue`)
    pulseOrigMmcssResponsiveness = null
  }

  // Restore MMCSS Tasks\\Games to safe defaults
  await runPS(`
    $base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games'
    Set-ItemProperty -Path $base -Name 'GPU Priority'          -Value 8     -Force -EA SilentlyContinue
    Set-ItemProperty -Path $base -Name 'Priority'              -Value 2     -Force -EA SilentlyContinue
    Set-ItemProperty -Path $base -Name 'Scheduling Category'   -Value 'Medium' -Force -EA SilentlyContinue
  `)

  // Restore power plan
  if (pulseOrigPowerPlan) {
    await runPS(`powercfg /setactive "${pulseOrigPowerPlan}"`)
    pulseOrigPowerPlan = null
  }

  // Restore DWM I/O + page priority to normal
  await runPS(`
    $dwm = Get-Process -Name dwm -EA SilentlyContinue | Select-Object -First 1
    if ($dwm) {
      try {
        Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class ProcPrio {
  [DllImport("ntdll.dll")] public static extern int NtSetInformationProcess(IntPtr h, int cls, ref int val, int len);
}
'@ -EA SilentlyContinue
        $ioPrio = 2   # IO_PRIORITY_HINT: 2 = Low (Windows default for DWM)
        $pgPrio = 5   # MEMORY_PRIORITY: 5 = Normal
        [ProcPrio]::NtSetInformationProcess($dwm.Handle, 21, [ref]$ioPrio, 4) | Out-Null
        [ProcPrio]::NtSetInformationProcess($dwm.Handle, 33, [ref]$pgPrio, 4) | Out-Null
      } catch {}
    }
  `, 8000)

  // Restore all process priorities to Normal (skip critical system procs)
  await runPS(`
    Get-Process -EA SilentlyContinue | Where-Object {
      $_.Name -notmatch '^(System|Idle|Registry|smss|csrss|wininit|lsass|services|svchost)$'
    } | ForEach-Object { try { $_.PriorityClass = 'Normal' } catch {} }
  `, 15000)

  send('◆ Pulse deactivated — all settings restored.', 'ok')
}

async function activatePulse(presetId, killBg, send) {
  const preset = PULSE_PRESETS[presetId]
  if (!preset) return { ok: false, error: 'Unknown preset' }
  if (pulseActivePreset) await pulseRestoreAll(() => {})

  send(`◆ Pulse activating — ${preset.name}`, 'head')
  pulseActivePreset = presetId
  lastUsedPulsePreset = presetId

  // ── 1. Timer resolution: request 0.5ms via NtSetTimerResolution ──────────────
  // 5000 = 0.5ms in 100ns units. Falls back silently if ntdll isn't available.
  await runPS(`
    try {
      Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class TimerRes { [DllImport("ntdll.dll")] public static extern int NtSetTimerResolution(uint r, bool set, out uint cur); }
'@
      $cur = [uint32]0
      [TimerRes]::NtSetTimerResolution(5000, $true, [ref]$cur) | Out-Null
      Write-Output "TIMER_OK"
    } catch { Write-Output "TIMER_SKIP" }
  `, 10000).then(r => {
    if (r.out.trim() === 'TIMER_OK') send('  Timer resolution → 0.5ms', 'ok')
    else send('  Timer resolution: skipped (ntdll unavailable)', 'info')
  })

  // ── 2. CPU core unparking ─────────────────────────────────────────────────────
  // Parked cores take ~1ms to wake — causes micro-stutters when game bursts load
  await runPS(`
    $parkGuid  = '0cc5b647-c1df-4637-891a-dec35c318583'
    $perfBoost = 'be337238-0d82-4146-a960-4f3749d470c7'
    $subPower  = '54533251-82be-4824-96c1-47b60b740d00'
    $schemes = powercfg /list | Select-String 'Power Scheme GUID' | ForEach-Object { ($_ -split ':\s+')[1].Split(' ')[0] }
    foreach ($s in $schemes) {
      powercfg /setacvalueindex  $s $subPower $parkGuid  0   2>$null
      powercfg /setdcvalueindex  $s $subPower $parkGuid  0   2>$null
      powercfg /setacvalueindex  $s $subPower $perfBoost 3   2>$null
    }
    powercfg /update-settings 2>$null
    Write-Output "UNPARK_OK"
  `, 20000).then(r => {
    if (r.out.includes('UNPARK_OK')) send('  CPU core parking disabled', 'ok')
  })

  // ── 3. Switch to High Performance power plan ─────────────────────────────────
  const planR = await runPS(`(powercfg /getactivescheme) -replace '.*GUID:\\s+(\\S+).*','$1'`)
  pulseOrigPowerPlan = planR.out.trim()
  await runPS(`powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c`).then(r => {
    if (r.ok) send('  Power plan → High Performance', 'ok')
    else send('  Power plan: already optimal or plan unavailable', 'info')
  })

  // ── 4. MMCSS boost ────────────────────────────────────────────────────────────
  // Save current value so we can restore it exactly
  const mmcssR = await runPS(`(Get-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile' -EA SilentlyContinue).SystemResponsiveness`)
  pulseOrigMmcssResponsiveness = parseInt(mmcssR.out.trim()) || 20
  await runPS(`
    $base = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
    Set-ItemProperty -Path $base -Name SystemResponsiveness -Value 10 -Force
    $gbase = "$base\\Tasks\\Games"
    Set-ItemProperty -Path $gbase -Name 'GPU Priority'        -Value 8      -Force -EA SilentlyContinue
    Set-ItemProperty -Path $gbase -Name 'Priority'            -Value 6      -Force -EA SilentlyContinue
    Set-ItemProperty -Path $gbase -Name 'Scheduling Category' -Value 'High' -Force -EA SilentlyContinue
  `)
  send('  MMCSS Games scheduling → High (SystemResponsiveness=10)', 'ok')

  // ── 5. Interrupt affinity — move GPU + top NIC DPCs off CPU 0 ────────────────
  // CPU 0 is the default Windows interrupt target; moving heavy devices to CPU 2
  // frees it for the game's main thread and OS scheduler.
  const coreCount = os.cpus().length
  if (coreCount >= 4) {
    await runPS(`
      $targetCpu = 2  # CPU 2 — leave 0 for OS scheduler, 1 for game main thread
      $mask = [uint32](1 -shl $targetCpu)
      $pciBase = 'HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\PCI'
      $netGuid = '{4d36e972-e325-11ce-bfc1-08002be10318}'
      $displayGuid = '{4d36e968-e325-11ce-bfc1-08002be10318}'
      $moved = 0
      Get-ChildItem $pciBase -EA SilentlyContinue | ForEach-Object {
        Get-ChildItem $_.PSPath -EA SilentlyContinue | ForEach-Object {
          $props = Get-ItemProperty -Path $_.PSPath -EA SilentlyContinue
          $cg = $props.ClassGUID
          # Target GPU and top NIC only (not Wi-Fi)
          $isDisplay = $cg -eq $displayGuid
          $isEthernet = $cg -eq $netGuid -and $props.Service -notmatch 'iwifi|netathr|bcmwl|Netwtw|IntelWifi|rt[l6]8|rtswlan'
          if ($isDisplay -or $isEthernet) {
            $affPath = "$($_.PSPath)\\Device Parameters\\Interrupt Management\\Affinity Policy"
            New-Item -Path $affPath -Force -EA SilentlyContinue | Out-Null
            Set-ItemProperty -Path $affPath -Name DevicePolicy     -Value 4 -Force -EA SilentlyContinue
            Set-ItemProperty -Path $affPath -Name AssignmentSetOverride -Value $mask -Force -EA SilentlyContinue
            $moved++
          }
        }
      }
      Write-Output "AFFINITY_OK:$moved"
    `, 20000).then(r => {
      const m = r.out.match(/AFFINITY_OK:(\d+)/)
      if (m) send(`  Interrupt affinity routed for ${m[1]} device(s) → CPU 2`, 'ok')
    })
  }

  // ── 6. Kill background processes ─────────────────────────────────────────────
  if (killBg) {
    let killed = 0
    for (const name of preset.killList) {
      const r = await runPS(`$p = Get-Process -Name '${name}' -EA SilentlyContinue; if ($p) { Stop-Process -Name '${name}' -Force -EA SilentlyContinue; Write-Output 'killed' }`)
      if (r.out.trim() === 'killed') { send(`  Stopped: ${name}`, 'ok'); killed++ }
    }
    if (killed === 0) send('  No background processes found to stop', 'info')
  }

  // ── 7. Game process priority ──────────────────────────────────────────────────
  const exePattern = preset.exe.replace('_GTAProcess', '*GTAProcess')
  const prioR = await runPS(`
    $p = Get-Process -EA SilentlyContinue | Where-Object { $_.Name -like '${exePattern}' } | Select-Object -First 1
    if ($p) {
      try { $p.PriorityClass = '${preset.priority}' } catch {}
      Write-Output "PRIO_SET:$($p.Id)"
    } else { Write-Output "PRIO_NOTFOUND" }
  `)
  if (prioR.out.startsWith('PRIO_SET')) send(`  Game priority → ${preset.priority} (${preset.name})`, 'ok')
  else send(`  Game process not found yet — priority will apply when game launches`, 'info')

  // ── 8. DWM deprioritisation ───────────────────────────────────────────────────
  // Lower DWM's I/O and memory page priority (NOT CPU priority — that would break the desktop).
  // This reduces memory bandwidth DWM steals from the GPU during fullscreen/borderless games.
  await runPS(`
    $dwm = Get-Process -Name dwm -EA SilentlyContinue | Select-Object -First 1
    if ($dwm) {
      Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class ProcPrio {
  [DllImport("ntdll.dll")] public static extern int NtSetInformationProcess(IntPtr h, int cls, ref int val, int len);
}
'@ -EA SilentlyContinue
      try {
        $handle = $dwm.Handle
        $ioPrio = 1   # IO_PRIORITY_HINT: 1 = VeryLow
        $pgPrio = 1   # MEMORY_PRIORITY: 1 = VeryLow
        [ProcPrio]::NtSetInformationProcess($handle, 21, [ref]$ioPrio, 4) | Out-Null  # ProcessIoPriority
        [ProcPrio]::NtSetInformationProcess($handle, 33, [ref]$pgPrio, 4) | Out-Null  # ProcessPagePriority
        Write-Output "DWM_OK"
      } catch { Write-Output "DWM_SKIP" }
    } else { Write-Output "DWM_SKIP" }
  `, 10000).then(r => {
    if (r.out.trim() === 'DWM_OK') send('  DWM I/O + page priority lowered (frees memory bandwidth for GPU)', 'ok')
    else send('  DWM deprioritisation: skipped', 'info')
  })

  // ── 9. Background process deprioritisation ────────────────────────────────────
  await runPS(`
    $gamePid = (Get-Process -EA SilentlyContinue | Where-Object { $_.Name -like '${exePattern}' } | Select-Object -First 1)?.Id
    Get-Process -EA SilentlyContinue | Where-Object {
      $_.Id -ne $gamePid -and
      $_.Name -notmatch '^(System|Idle|Registry|smss|csrss|wininit|lsass|services|svchost|MsMpEng|audiodg|dwm)$'
    } | ForEach-Object { try { $_.PriorityClass = 'BelowNormal' } catch {} }
  `, 15000)
  send(`  Background processes deprioritised`, 'ok')

  send(`◆ Pulse active — ${preset.name}`, 'head')
  discordPulseGame = preset.name
  updateDiscordPresence()
  sessionPulseUses++
  webhookPulseActivated(preset.name, [
    'Timer resolution → 0.5ms (NtSetTimerResolution)',
    'CPU core parking disabled',
    'Power plan → High Performance',
    'MMCSS Games scheduling → High (SystemResponsiveness=10)',
    `Interrupt affinity routed → CPU 2 (${os.cpus().length >= 4 ? 'applied' : 'skipped — <4 cores'})`,
    'DWM I/O + page priority lowered',
    `Game priority → ${preset.priority} (${preset.name})`,
    killBg ? `Background apps killed (${preset.killList.join(', ')})` : 'Background kill: skipped',
  ])
  mainWindow?.webContents.send('pulse-tick', { preset: presetId, active: true })
  updateTrayMenu()
  return { ok: true }
}

async function deactivatePulse(send) {
  pulseActivePreset = null
  discordPulseGame = null
  updateDiscordPresence()
  await pulseRestoreAll(send)
  mainWindow?.webContents.send('pulse-tick', { active: false })
  updateTrayMenu()
  return { ok: true }
}

ipcMain.handle('pulse-start', async (_, { presetId, killBg }) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  return activatePulse(presetId, killBg, send)
})

ipcMain.handle('pulse-stop', async () => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  autoPulseTriggeredPreset = null  // clear auto-pulse state on manual stop too
  return deactivatePulse(send)
})

ipcMain.handle('pulse-get-presets', () => {
  return Object.entries(PULSE_PRESETS).map(([id, p]) => ({ id, name: p.name, exe: p.exe }))
})

ipcMain.handle('pulse-detect-game', async () => {
  const r = await runPS(`Get-Process | Select-Object -ExpandProperty Name`)
  const running = r.out.toLowerCase()
  for (const [id, p] of Object.entries(PULSE_PRESETS)) {
    const pattern = p.exe.toLowerCase()
    const lines = running.split('\n').map(l => l.trim())
    const match = pattern.includes('_gtaprocess')
      ? lines.some(l => l.endsWith('_gtaprocess'))
      : pattern.endsWith('*')
        ? lines.some(l => l.startsWith(pattern.slice(0, -1)))
        : lines.some(l => l === pattern)
    if (match) return { found: true, id, name: p.name }
  }
  return { found: false }
})

// ─── Startup Manager ──────────────────────────────────────────────────────────
ipcMain.handle('get-startup-items', async () => {
  const r = await runPS(`
    $items = @()
    # Registry HKCU Run
    $p = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    if (Test-Path $p) {
      (Get-ItemProperty $p).PSObject.Properties | Where-Object {$_.Name -notmatch '^PS'} | ForEach-Object {
        $items += [PSCustomObject]@{Name=$_.Name;Command=$_.Value;Source='HKCU Registry';Enabled=$true;Impact='Low';Key=$p}
      }
    }
    # Registry HKLM Run
    $p2 = 'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    if (Test-Path $p2) {
      (Get-ItemProperty $p2).PSObject.Properties | Where-Object {$_.Name -notmatch '^PS'} | ForEach-Object {
        $items += [PSCustomObject]@{Name=$_.Name;Command=$_.Value;Source='HKLM Registry';Enabled=$true;Impact='Medium';Key=$p2}
      }
    }
    # Task Scheduler startup tasks
    Get-ScheduledTask | Where-Object {$_.Triggers.CimClass.CimClassName -contains 'MSFT_TaskLogonTrigger' -or $_.Triggers.CimClass.CimClassName -contains 'MSFT_TaskBootTrigger'} |
    Select-Object TaskName,TaskPath,@{N='State';E={$_.State}} | ForEach-Object {
      $items += [PSCustomObject]@{Name=$_.TaskName;Command=$_.TaskPath;Source='Task Scheduler';Enabled=($_.State -ne 'Disabled');Impact='Medium';Key=$_.TaskName}
    }
    $items | ConvertTo-Json -Depth 2
  `, 20000)
  try {
    const data = JSON.parse(r.out)
    return Array.isArray(data) ? data : [data]
  } catch { return [] }
})

ipcMain.handle('toggle-startup-item', async (_, { name, source, key, enable }) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  if (source === 'Task Scheduler') {
    const cmd = enable ? `Enable-ScheduledTask -TaskName '${name}' -ErrorAction SilentlyContinue` : `Disable-ScheduledTask -TaskName '${name}' -ErrorAction SilentlyContinue`
    await runPS(cmd)
    send(`${enable ? 'Enabled' : 'Disabled'} task: ${name}`, 'ok')
  } else {
    if (!enable) {
      // Move to disabled key
      const disabledKey = key.replace('\\Run', '\\Run\\AutorunsDisabled')
      await runPS(`
        $val = (Get-ItemProperty -Path '${key}' -Name '${name}' -EA SilentlyContinue).'${name}'
        if ($val) {
          New-Item -Path '${disabledKey}' -Force -EA SilentlyContinue | Out-Null
          New-ItemProperty -Path '${disabledKey}' -Name '${name}' -Value $val -Force
          Remove-ItemProperty -Path '${key}' -Name '${name}' -Force -EA SilentlyContinue
        }
      `)
    } else {
      const disabledKey = key.replace('\\Run', '\\Run\\AutorunsDisabled')
      await runPS(`
        $val = (Get-ItemProperty -Path '${disabledKey}' -Name '${name}' -EA SilentlyContinue).'${name}'
        if ($val) {
          New-ItemProperty -Path '${key}' -Name '${name}' -Value $val -Force
          Remove-ItemProperty -Path '${disabledKey}' -Name '${name}' -Force -EA SilentlyContinue
        }
      `)
    }
    send(`${enable ? 'Enabled' : 'Disabled'} startup item: ${name}`, 'ok')
  }
  return { ok: true }
})

// ─── Network Ping Test ────────────────────────────────────────────────────────
ipcMain.handle('run-ping-tests', async () => {
  mainWindow?.webContents.send('log', { msg: '◆ Running ping tests…', level: 'head', ts: new Date().toLocaleTimeString() })

  const targets = [
    { name: 'Cloudflare DNS',    host: '1.1.1.1',          category: 'DNS' },
    { name: 'Google DNS',        host: '8.8.8.8',           category: 'DNS' },
    { name: 'Riot Games (EU)',   host: 'euc1.lol.gamesvc.net', category: 'Valorant/LoL' },
    { name: 'Riot Games (NA)',   host: 'nac.lol.gamesvc.net',  category: 'Valorant/LoL' },
    { name: 'Steam Content',     host: 'content1.st.dl.eccdnx.com', category: 'Steam' },
    { name: 'EA / Origin',       host: 'ea.com',            category: 'EA/Apex' },
    { name: 'Activision',        host: 'activision.com',    category: 'Warzone/CoD' },
    { name: 'Epic Games',        host: 'epicgames.com',     category: 'Fortnite' },
    { name: 'Rockstar Games',    host: 'ros.rockstargames.com', category: 'GTA/FiveM' },
    { name: 'FiveM CFX',         host: 'runtime.fivem.net', category: 'FiveM' },
    { name: 'Google (General)',  host: '8.8.4.4',           category: 'General' },
    { name: 'Cloudflare (Alt)',  host: '1.0.0.1',           category: 'DNS' },
  ]

  const results = []
  for (const t of targets) {
    const r = await runPS(`
      $pings = 1..4 | ForEach-Object {
        try {
          $p = New-Object System.Net.NetworkInformation.Ping
          $reply = $p.Send('${t.host}', 2000)
          if ($reply.Status -eq 'Success') { $reply.RoundtripTime } else { 9999 }
        } catch { 9999 }
      }
      $valid = $pings | Where-Object {$_ -lt 9999}
      if ($valid) {
        $avg = [math]::Round(($valid | Measure-Object -Average).Average)
        $min = ($valid | Measure-Object -Minimum).Minimum
        Write-Output "$avg|$min|$($valid.Count)"
      } else { Write-Output "timeout|timeout|0" }
    `, 12000)
    const parts = r.out.trim().split('|')
    const avg = parts[0] === 'timeout' ? null : parseInt(parts[0])
    const min = parts[1] === 'timeout' ? null : parseInt(parts[1])
    const count = parseInt(parts[2]) || 0
    results.push({ ...t, avg, min, packetLoss: Math.round((4 - count) / 4 * 100) })
    mainWindow?.webContents.send('ping-result', { ...t, avg, min, packetLoss: Math.round((4 - count) / 4 * 100) })
  }
  return results
})

// ─── Game Detection ───────────────────────────────────────────────────────────
ipcMain.handle('detect-games', async () => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
  send('Scanning for installed games…', 'info')

  const GAME_DEFINITIONS = [
    { id:'fortnite',      name:'Fortnite',              exe:'FortniteClient-Win64-Shipping.exe', searchNames:['Fortnite','FortniteGame'] },
    { id:'gtav',          name:'GTA V',                 exe:'GTA5.exe',              searchNames:['Grand Theft Auto V','GTAV'] },
    { id:'valorant',      name:'Valorant',              exe:'VALORANT-Win64-Shipping.exe', searchNames:['VALORANT','Riot Games\\VALORANT'] },
    { id:'arma-reforger', name:'Arma Reforger',         exe:'ArmaReforger.exe',      searchNames:['Arma Reforger'] },
    { id:'minecraft',     name:'Minecraft',             exe:'javaw.exe',             searchNames:['Minecraft Launcher','.minecraft'] },
    { id:'cs2',           name:'Counter-Strike 2',      exe:'cs2.exe',               searchNames:['Counter-Strike Global Offensive','cs2'] },
    { id:'apex',          name:'Apex Legends',          exe:'r5apex.exe',            searchNames:['Apex Legends'] },
    { id:'warzone',       name:'Warzone',               exe:'cod.exe',               searchNames:['Call of Duty','Modern Warfare','Warzone'] },
    { id:'rust',          name:'Rust',                  exe:'RustClient.exe',        searchNames:['Rust'] },
    { id:'the-finals',    name:'THE FINALS',            exe:'FINALS-Win64-Shipping.exe', searchNames:['THE FINALS','Discovery'] },
  ]

  // Common install base paths
  const steamPath = await runPS('(Get-ItemProperty "HKLM:\\SOFTWARE\\WOW6432Node\\Valve\\Steam" -ErrorAction SilentlyContinue).InstallPath')
  const epicPath  = await runPS('(Get-ItemProperty "HKLM:\\SOFTWARE\\WOW6432Node\\EpicGames\\EpicGamesLauncher" -ErrorAction SilentlyContinue).AppDataPath')
  const riotPath  = await runPS('(Get-ItemProperty "HKLM:\\SOFTWARE\\WOW6432Node\\Riot Games\\VALORANT" -ErrorAction SilentlyContinue).InstallLocation')

  const steamBase = steamPath.out.trim()
  const searchBases = [
    steamBase ? path.join(steamBase, 'steamapps', 'common') : '',
    'C:\\Program Files\\Steam\\steamapps\\common',
    'C:\\Program Files (x86)\\Steam\\steamapps\\common',
    'C:\\Program Files\\Epic Games',
    'C:\\Program Files (x86)\\Epic Games',
    'C:\\Program Files\\Riot Games',
    'C:\\Program Files\\EA Games',
    'C:\\Program Files\\Rockstar Games',
    'C:\\XboxGames',
    path.join(os.homedir(), 'AppData', 'Local', 'FortniteGame'),
    epicPath.out.trim() || '',
  ].filter(Boolean)

  const found = []
  for (const game of GAME_DEFINITIONS) {
    let installPath = null
    let lastPlayed = null

    // Search install bases
    for (const base of searchBases) {
      if (!base || !fs.existsSync(base)) continue
      for (const name of game.searchNames) {
        const candidate = path.join(base, name)
        if (fs.existsSync(candidate)) {
          installPath = candidate
          try {
            const stats = fs.statSync(candidate)
            lastPlayed = stats.mtime.toLocaleDateString()
          } catch {}
          break
        }
      }
      if (installPath) break
    }

    // Registry check for any exe path
    if (!installPath) {
      const regCheck = await runPS(`
        @('HKLM:\\SOFTWARE','HKLM:\\SOFTWARE\\WOW6432Node','HKCU:\\SOFTWARE') | ForEach-Object {
          $base = $_
          ${game.searchNames.map(n => `
          try { $p = Get-ItemProperty "$base\\${n}" -EA SilentlyContinue; if ($p.InstallLocation -or $p.InstallDir) { Write-Output ($p.InstallLocation ?? $p.InstallDir); return } } catch {}
          `).join('')}
        }
      `)
      if (regCheck.out.trim()) { installPath = regCheck.out.trim().split('\n')[0].trim() }
    }

    // Check if exe is currently in process list (game is running)
    const isRunning = await runPS(`if (Get-Process -Name '${game.exe.replace('.exe','')}' -EA SilentlyContinue) { Write-Output 'yes' } else { Write-Output 'no' }`)

    found.push({
      ...game,
      installed: !!installPath,
      installPath: installPath || null,
      lastPlayed: lastPlayed || 'Unknown',
      isRunning: isRunning.out.trim() === 'yes'
    })
  }

  const installedCount = found.filter(g => g.installed).length
  send(`Game scan complete — ${installedCount} installed games detected.`, 'ok')
  return found
})

// ─── Per-Game Auto Profile (process watcher) ──────────────────────────────────
let gameWatcherInterval = null
let activeGameProfile = null

ipcMain.handle('start-game-watcher', async () => {
  if (gameWatcherInterval) return { ok: true, status: 'already running' }

  mainWindow?.webContents.send('log', { msg: '◆ Game Watcher started — monitoring for game launches…', level: 'ok', ts: new Date().toLocaleTimeString() })

  gameWatcherInterval = setInterval(async () => {
    const r = await runPS("Get-Process | Select-Object -ExpandProperty Name", 5000)
    const running = new Set(r.out.split('\n').map(l => l.trim().toLowerCase()))

    // Detect by matching directly against PULSE_PRESETS exe names
    let detectedGame = null
    for (const [presetId, preset] of Object.entries(PULSE_PRESETS)) {
      const exeLower = preset.exe.toLowerCase()
      const match = exeLower.includes('_gtaprocess')
        ? [...running].some(p => p.endsWith('_gtaprocess'))
        : exeLower.endsWith('*')
          ? [...running].some(p => p.startsWith(exeLower.slice(0, -1)))
          : running.has(exeLower)
      if (match) { detectedGame = presetId; break }
    }

    if (detectedGame && activeGameProfile !== detectedGame) {
      // New game detected — apply tweaks
      activeGameProfile = detectedGame
      const presetName = PULSE_PRESETS[detectedGame].name
      mainWindow?.webContents.send('game-watcher-event', { event: 'game-started', gameId: detectedGame })
      mainWindow?.webContents.send('log', { msg: `🎮 Game detected: ${presetName} — applying optimizations…`, level: 'head', ts: new Date().toLocaleTimeString() })

      // Update Discord presence to show the active game
      discordActiveGame = detectedGame
      updateDiscordPresence()

      const mmcssBase = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
      await runPS(`Set-ItemProperty -Path "${mmcssBase}" -Name SystemResponsiveness -Value 10 -Force`)
      await runPS(`Set-ItemProperty -Path "${mmcssBase}\\Tasks\\Games" -Name "GPU Priority" -Value 8 -Force`)
      await runPS('Set-ItemProperty -Path "HKCU:\\System\\GameConfigStore" -Name GameDVR_Enabled -Value 0 -Force')
      mainWindow?.webContents.send('log', { msg: `✓ Auto-profile applied for ${presetName}`, level: 'ok', ts: new Date().toLocaleTimeString() })

      // Auto-Pulse: activate if enabled (preset already matched — use it directly)
      const currentSettings = loadSettings()
      if (currentSettings.autoPulseEnabled && !pulseActivePreset) {
        autoPulseTriggeredPreset = detectedGame
        const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
        send(`◆ Auto-Pulse: activating for ${presetName}…`, 'head')
        mainWindow?.webContents.send('game-watcher-event', { event: 'auto-pulse-start', gameId: detectedGame, presetId: detectedGame })
        activatePulse(detectedGame, false, send).catch(() => { autoPulseTriggeredPreset = null })
      }

    } else if (!detectedGame && activeGameProfile) {
      // Game closed — restore
      const prev = activeGameProfile
      activeGameProfile = null
      mainWindow?.webContents.send('game-watcher-event', { event: 'game-closed', gameId: prev })
      mainWindow?.webContents.send('log', { msg: `Game closed: ${prev} — restoring defaults…`, level: 'info', ts: new Date().toLocaleTimeString() })

      // Clear game from Discord presence
      discordActiveGame = null
      updateDiscordPresence()

      const mmcssBase = 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile'
      await runPS(`Set-ItemProperty -Path "${mmcssBase}" -Name SystemResponsiveness -Value 20 -Force`)
      mainWindow?.webContents.send('log', { msg: '✓ Defaults restored.', level: 'ok', ts: new Date().toLocaleTimeString() })

      // Auto-Pulse: deactivate if it was auto-triggered
      if (autoPulseTriggeredPreset) {
        const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
        send(`◆ Auto-Pulse: deactivating (${prev} closed)`, 'info')
        autoPulseTriggeredPreset = null
        mainWindow?.webContents.send('game-watcher-event', { event: 'auto-pulse-stop', gameId: prev })
        deactivatePulse(send).catch(() => {})
      }
    }
  }, 5000)

  return { ok: true, status: 'started' }
})

ipcMain.handle('stop-game-watcher', async () => {
  if (gameWatcherInterval) { clearInterval(gameWatcherInterval); gameWatcherInterval = null }
  activeGameProfile = null
  discordActiveGame = null
  updateDiscordPresence()
  // If auto-pulse was active, deactivate it
  if (autoPulseTriggeredPreset) {
    autoPulseTriggeredPreset = null
    const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })
    deactivatePulse(send).catch(() => {})
  }
  mainWindow?.webContents.send('log', { msg: 'Game Watcher stopped.', level: 'info', ts: new Date().toLocaleTimeString() })
  return { ok: true }
})

ipcMain.handle('get-game-watcher-status', () => ({ running: !!gameWatcherInterval, activeGame: activeGameProfile }))

// ─── Tweak Changelog ──────────────────────────────────────────────────────────
const CHANGELOG_PATH = path.join(app.getPath('userData'), 'jt_changelog.json')

function loadChangelog() {
  try { if (fs.existsSync(CHANGELOG_PATH)) return JSON.parse(fs.readFileSync(CHANGELOG_PATH, 'utf8')) } catch {}
  return []
}
function saveChangelog(log) {
  try { fs.writeFileSync(CHANGELOG_PATH, JSON.stringify(log, null, 2)) } catch {}
}

ipcMain.handle('get-changelog', () => loadChangelog())
ipcMain.handle('add-changelog-entry', (_, entry) => {
  const log = loadChangelog()
  log.unshift({ ...entry, ts: new Date().toISOString() })
  if (log.length > 200) log.splice(200)
  saveChangelog(log)
  return { ok: true }
})
ipcMain.handle('clear-changelog', () => { saveChangelog([]); return { ok: true } })

// ─── Tweak Scheduler ─────────────────────────────────────────────────────────
ipcMain.handle('schedule-tweak', async (_, { taskName, schedule, action }) => {
  const send = (msg, level) => mainWindow?.webContents.send('log', { msg, level, ts: new Date().toLocaleTimeString() })

  const psCommands = {
    'clean-temp':    'Remove-Item -Path $env:TEMP\\* -Recurse -Force -EA SilentlyContinue; Remove-Item -Path C:\\Windows\\Temp\\* -Recurse -Force -EA SilentlyContinue; Remove-Item -Path C:\\Windows\\Prefetch\\* -Force -EA SilentlyContinue',
    'flush-dns':     'ipconfig /flushdns',
    'clean-nvidia':  'Remove-Item -Path "$env:LOCALAPPDATA\\NVIDIA\\DXCache\\*" -Recurse -Force -EA SilentlyContinue; Remove-Item -Path "$env:LOCALAPPDATA\\NVIDIA\\GLCache\\*" -Recurse -Force -EA SilentlyContinue; Remove-Item -Path "$env:LOCALAPPDATA\\Temp\\NVIDIA Corporation\\NV_Cache\\*" -Recurse -Force -EA SilentlyContinue',
    'defrag-c':      'Optimize-Volume -DriveLetter C -Defrag -Verbose',
  }
  const psCommand = psCommands[action] || action

  // Encode as UTF-16 LE Base64 so -EncodedCommand handles all quoting automatically
  const encodedCommand = Buffer.from(psCommand, 'utf16le').toString('base64')

  let trigger = ''
  if (schedule === 'startup')    trigger = 'New-ScheduledTaskTrigger -AtStartup'
  else if (schedule === 'daily-3am')  trigger = 'New-ScheduledTaskTrigger -Daily -At "3:00AM"'
  else if (schedule === 'weekly-sun') trigger = 'New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "3:00AM"'
  else if (schedule === 'logon')      trigger = 'New-ScheduledTaskTrigger -AtLogon'

  const r = await runPS(`
    try {
      $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NonInteractive -WindowStyle Hidden -EncodedCommand ${encodedCommand}'
      $trigger   = ${trigger}
      $settings  = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable:$false -StartWhenAvailable
      $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
      Register-ScheduledTask -TaskName 'JT_${taskName.replace(/\W/g, '_')}' -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force -ErrorAction Stop | Out-Null
      Write-Output 'ok'
    } catch {
      Write-Error $_.Exception.Message
    }
  `)

  if (r.out.trim() === 'ok') {
    send(`✓ Scheduled task created: JT_${taskName} (${schedule})`, 'ok')
    return { ok: true }
  }
  send(`Failed to create task: ${r.err || r.out}`, 'err')
  return { ok: false, error: r.err || r.out }
})

ipcMain.handle('list-scheduled-tweaks', async () => {
  const r = await runPS('Get-ScheduledTask | Where-Object {$_.TaskName -like "JT_*"} | Select-Object TaskName,@{N="State";E={$_.State}} | ConvertTo-Json')
  try { const d = JSON.parse(r.out); return Array.isArray(d) ? d : [d] } catch { return [] }
})

ipcMain.handle('delete-scheduled-tweak', async (_, taskName) => {
  await runPS(`Unregister-ScheduledTask -TaskName '${taskName}' -Confirm:$false -ErrorAction SilentlyContinue`)
  return { ok: true }
})

// ─── PC Health Score ──────────────────────────────────────────────────────────
ipcMain.handle('get-temp-size', async () => {
  const paths = [
    process.env.TEMP,
    'C:\\Windows\\Temp',
    'C:\\Windows\\Prefetch',
    'C:\\Windows\\SoftwareDistribution\\Download',
  ].filter(Boolean)

  const r = await runPS(`
    $total = 0
    $paths = @(${paths.map(p => `'${p}'`).join(',')})
    foreach ($p in $paths) {
      if (Test-Path $p) {
        $total += (Get-ChildItem -Path $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
      }
    }
    Write-Output $total
  `)
  const bytes = parseInt(r.out?.trim()) || 0
  const mb = bytes / 1048576
  let label, pct
  if (mb < 100)       { label = `${Math.round(mb)} MB`;  pct = Math.round((mb / 500) * 100) }
  else if (mb < 1024) { label = `${Math.round(mb)} MB`;  pct = Math.round((mb / 2048) * 100) }
  else                { label = `${(mb / 1024).toFixed(1)} GB`; pct = Math.min(100, Math.round((mb / 4096) * 100)) }
  return { bytes, mb: Math.round(mb), label, pct }
})

ipcMain.handle('get-health-score', async (_, settings) => {
  const scores = {}

  // 1. Temp files size (0-20 points)
  const tempR = await runPS(`
    $size = 0
    @("$env:TEMP","C:\\Windows\\Temp","C:\\Windows\\Prefetch") | ForEach-Object {
      if (Test-Path $_) { $size += (Get-ChildItem $_ -Recurse -EA SilentlyContinue | Measure-Object -Property Length -Sum).Sum }
    }
    Write-Output ([math]::Round($size/1MB))
  `)
  const tempMB = parseInt(tempR.out.trim()) || 0
  scores.temp = tempMB < 100 ? 20 : tempMB < 500 ? 15 : tempMB < 2000 ? 8 : 2

  // 2. Startup items count (0-15 points)
  const startupR = await runPS(`
    $count = 0
    $p = 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    if (Test-Path $p) { $count += ((Get-ItemProperty $p).PSObject.Properties | Where-Object {$_.Name -notmatch '^PS'}).Count }
    $p2 = 'HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    if (Test-Path $p2) { $count += ((Get-ItemProperty $p2).PSObject.Properties | Where-Object {$_.Name -notmatch '^PS'}).Count }
    Write-Output $count
  `)
  const startupCount = parseInt(startupR.out.trim()) || 0
  scores.startup = startupCount <= 4 ? 15 : startupCount <= 8 ? 10 : startupCount <= 14 ? 5 : 2

  // 3. Disk health (0-20 points)
  const diskR = await runPS(`
    try {
      $disk = Get-PhysicalDisk | Where-Object {$_.MediaType -ne 'Unspecified'} | Select-Object -First 1
      $health = $disk.HealthStatus
      Write-Output $health
    } catch { Write-Output 'Unknown' }
  `)
  const diskHealth = diskR.out.trim()
  scores.disk = diskHealth === 'Healthy' ? 20 : diskHealth === 'Warning' ? 8 : diskHealth === 'Unhealthy' ? 0 : 12

  // 4. Tweaks applied (0-30 points)
  const appliedTweaks = Object.entries(settings || {}).filter(([k,v]) => k.startsWith('tweak_') && v === 'applied').length
  const RECOMMENDED_TWEAKS = ['game-dvr','visual-effects','mmcss','gpu-hwsch','ntfs','qos-reserve','net-throttling','tcp-stack','disable-netbios','disable-wpad']
  const recommendedApplied = RECOMMENDED_TWEAKS.filter(id => settings?.[`tweak_${id}`] === 'applied').length
  scores.tweaks = Math.min(30, Math.round((recommendedApplied / RECOMMENDED_TWEAKS.length) * 30))

  // 5. Driver freshness (0-15 points) — check if GPU driver is recent (within 6 months)
  const driverR = await runPS(`
    try {
      $d = Get-CimInstance Win32_PnPSignedDriver | Where-Object {$_.DeviceClass -eq 'Display'} | Select-Object -First 1
      if ($d.DriverDate) { Write-Output $d.DriverDate.ToString('yyyy-MM-dd') } else { Write-Output 'unknown' }
    } catch { Write-Output 'unknown' }
  `)
  const driverDate = driverR.out.trim()
  if (driverDate !== 'unknown') {
    const age = (Date.now() - new Date(driverDate)) / (1000 * 60 * 60 * 24)
    scores.drivers = age < 90 ? 15 : age < 180 ? 12 : age < 365 ? 7 : 3
  } else { scores.drivers = 8 }

  const total = Object.values(scores).reduce((a, b) => a + b, 0)

  return {
    total: Math.min(100, total),
    breakdown: {
      temp:    { score: scores.temp,    max: 20, label: 'Temp Files',     detail: `${tempMB} MB of temp files` },
      startup: { score: scores.startup, max: 15, label: 'Startup Items',  detail: `${startupCount} startup items` },
      disk:    { score: scores.disk,    max: 20, label: 'Disk Health',    detail: diskHealth },
      tweaks:  { score: scores.tweaks,  max: 30, label: 'Tweaks Applied', detail: `${recommendedApplied}/${RECOMMENDED_TWEAKS.length} recommended` },
      drivers: { score: scores.drivers, max: 15, label: 'Driver Freshness', detail: driverDate !== 'unknown' ? `Last updated ${driverDate}` : 'Unknown' },
    },
    tempMB, startupCount, diskHealth,
    appliedTweaks, driverDate
  }
})
