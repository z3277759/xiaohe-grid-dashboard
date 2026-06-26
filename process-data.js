/**
 * 数据处理脚本 - 读取所有Excel文件，生成可视化所需的JSON数据
 * 日看板：使用6月5日数据
 * 月看板：汇总6月1日-5日数据
 * 运行: node process-data.js
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const DATA_DIR = __dirname;
const OUTPUT_FILE = path.join(DATA_DIR, 'data.json');

// ===================== 辅助函数 =====================
function serialToDate(serial) {
    if (typeof serial === 'string') return serial.substring(0, 10);
    if (typeof serial === 'number') {
        return new Date((serial - 25569) * 86400 * 1000).toISOString().substring(0, 10);
    }
    return String(serial);
}

function extractDate(dt) {
    if (!dt) return '';
    return String(dt).substring(0, 10);
}

function normalizeOrgName(name) {
    if (!name) return '';
    return String(name).replace(/^\s*\[\d+\]\s*/, '').trim();
}

// ===================== 读取渠道归属 =====================
console.log('=== 读取渠道归属 ===');
const chMap = XLSX.readFile(path.join(DATA_DIR, '渠道归属.xlsx'));
const chSheet = chMap.Sheets['Sheet1'];
const chData = XLSX.utils.sheet_to_json(chSheet, { header: 1, defval: '' });

const orgGridMap = {};
const orgLineMap = {};
const allGrids = new Set();
const allLines = new Set();

for (let i = 1; i < chData.length; i++) {
    const [ch, grid, line] = chData[i];
    if (!ch) continue;
    const norm = normalizeOrgName(ch);
    if (grid) { orgGridMap[norm] = grid; allGrids.add(grid); }
    if (line) { orgLineMap[norm] = line; allLines.add(line); }
}
console.log('网格数:', allGrids.size, '\n产线数:', allLines.size);

// ===================== 读取函数 =====================
/** 读取 新入网/大套餐/FTTR 格式（9列：受理组织=2，受理时间=6，工单=7） */
function readOrders(filePath) {
    if (!fs.existsSync(filePath)) return [];
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const records = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[2]) continue;
        records.push({
            org: normalizeOrgName(row[2]),
            date: serialToDate(row[6]),
            orderId: String(row[7] || ''),
        });
    }
    return records;
}

/** 读取 宽带/携入 格式（20列：受理组织=8，受理时间=0，竣工=3，撤单=4） */
function readBB(filePath) {
    if (!fs.existsSync(filePath)) return [];
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const records = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[8]) continue;
        records.push({
            org: normalizeOrgName(row[8]),
            date: serialToDate(row[0]),
            completed: String(row[3] || '').trim() === '已竣工',
            cancelled: String(row[4] || '').trim().includes('撤单'),
        });
    }
    return records;
}

// ===================== 读取日数据（6月5日） =====================
console.log('\n=== 读取日数据 (6月5日) ===');
function readDay(day) {
    const prefix = `6月${day}`;
    return {
        xinruwang: readOrders(path.join(DATA_DIR, `${prefix}新入网.xlsx`)),
        dataitaocan: readOrders(path.join(DATA_DIR, `${prefix}大套餐.xlsx`)),
        fttr: readOrders(path.join(DATA_DIR, `${prefix}fttr.xlsx`)),
        broadband: readBB(path.join(DATA_DIR, `${prefix}宽带.xlsx`)),
        xielu: readBB(path.join(DATA_DIR, `${prefix}携入.xlsx`)),
    };
}

const dailyData = readDay(5);
console.log('  新入网:', dailyData.xinruwang.length, ' 大套餐:', dailyData.dataitaocan.length, ' FTTR:', dailyData.fttr.length,
    ' 宽带:', dailyData.broadband.filter(r=>r.completed).length, ' 携入:', dailyData.xielu.filter(r=>r.completed).length);

// ===================== 读取月数据（6月1日-5日） =====================
console.log('\n=== 读取月数据 (6月1日-5日) ===');
const monthData = { xinruwang: [], dataitaocan: [], fttr: [], broadband: [], xielu: [] };
for (let day = 1; day <= 5; day++) {
    const dd = readDay(day);
    Object.keys(monthData).forEach(k => monthData[k].push(...dd[k]));
}
console.log('  新入网:', monthData.xinruwang.length, ' 大套餐:', monthData.dataitaocan.length, ' FTTR:', monthData.fttr.length,
    ' 宽带:', monthData.broadband.filter(r=>r.completed).length, ' 携入:', monthData.xielu.filter(r=>r.completed).length);

// ===================== 汇总函数 =====================
const IKEYS = ['xinruwang', 'dataitaocan', 'fttr', 'broadband', 'xielu'];
const INAMES = { xinruwang: '新入网', dataitaocan: '大套餐', fttr: 'FTTR', broadband: '宽带新增', xielu: '携转' };

function countByOrg(records, opts = {}) {
    const counts = {};
    for (const r of records) {
        if (opts.checkCompleted && !r.completed) continue;
        if (opts.checkCancelled && r.cancelled) continue;
        counts[r.org] = (counts[r.org] || 0) + 1;
    }
    return counts;
}

function computeStats(rawData, label) {
    const xinruwangCounts = countByOrg(rawData.xinruwang);
    const dataitaocanCounts = countByOrg(rawData.dataitaocan);
    const fttrCounts = countByOrg(rawData.fttr);
    const broadbandCounts = countByOrg(rawData.broadband, { checkCompleted: true });
    const xieluCounts = countByOrg(rawData.xielu, { checkCompleted: true });
    const allCountMaps = { xinruwang: xinruwangCounts, dataitaocan: dataitaocanCounts, fttr: fttrCounts, broadband: broadbandCounts, xielu: xieluCounts };

    // 网格汇总
    const gridData = {};
    const gridOrgDetails = {};
    for (const [org, grid] of Object.entries(orgGridMap)) {
        if (!gridData[grid]) { gridData[grid] = {}; gridOrgDetails[grid] = {}; IKEYS.forEach(ind => gridData[grid][ind] = 0); }
        IKEYS.forEach(ind => {
            const count = allCountMaps[ind][org] || 0;
            gridData[grid][ind] += count;
            if (!gridOrgDetails[grid][org]) { gridOrgDetails[grid][org] = {}; IKEYS.forEach(i => gridOrgDetails[grid][org][i] = 0); }
            gridOrgDetails[grid][org][ind] += count;
        });
    }

    // 产线汇总
    const lineData = {};
    for (const [org, line] of Object.entries(orgLineMap)) {
        if (!line) continue;
        if (!lineData[line]) { lineData[line] = {}; IKEYS.forEach(ind => lineData[line][ind] = 0); }
        IKEYS.forEach(ind => lineData[line][ind] += allCountMaps[ind][org] || 0);
    }

    // 网格×产线
    const gridLineData = {};
    for (const [org, grid] of Object.entries(orgGridMap)) {
        const line = orgLineMap[org] || '未分类';
        const key = grid + '|' + line;
        if (!gridLineData[key]) { gridLineData[key] = { grid, line, orgs: new Set() }; IKEYS.forEach(ind => gridLineData[key][ind] = 0); }
        IKEYS.forEach(ind => { gridLineData[key][ind] += allCountMaps[ind][org] || 0; });
        gridLineData[key].orgs.add(org);
    }

    // 总览
    const overview = {};
    IKEYS.forEach(ind => {
        overview[ind] = Object.values(allCountMaps[ind]).reduce((s, v) => s + v, 0);
    });
    console.log(`${label} 总览:`, overview);

    return {
        overview,
        gridData: Object.entries(gridData).map(([grid, vals]) => ({ grid, ...vals, total: IKEYS.reduce((s, ind) => s + (vals[ind] || 0), 0) })).sort((a, b) => b.total - a.total),
        lineData: Object.entries(lineData).map(([line, vals]) => ({ line, ...vals, total: IKEYS.reduce((s, ind) => s + (vals[ind] || 0), 0) })).sort((a, b) => b.total - a.total),
        gridLineData: Object.entries(gridLineData).map(([k, val]) => ({ grid: val.grid, line: val.line, orgCount: val.orgs.size, orgs: [...val.orgs], ...IKEYS.reduce((o, ind) => (o[ind] = val[ind], o), {}), total: IKEYS.reduce((s, ind) => s + (val[ind] || 0), 0) })).sort((a, b) => a.grid.localeCompare(b.grid) || a.line.localeCompare(b.line)),
        orgDetails: Object.fromEntries(Object.entries(gridOrgDetails).map(([grid, orgs]) => [grid, Object.entries(orgs).map(([org, vals]) => ({ org, ...vals, total: IKEYS.reduce((s, ind) => s + (vals[ind] || 0), 0) })).sort((a, b) => b.total - a.total)])),
    };
}

const daily = computeStats(dailyData, '日数据(6/5)');
const monthly = computeStats(monthData, '月数据(6/1-6/5)');

// ===================== 输出JSON =====================
const output = {
    meta: {
        generatedAt: new Date().toISOString(),
        indicators: IKEYS,
        indicatorNames: INAMES,
        grids: [...allGrids].sort(),
        lines: [...allLines].sort().filter(Boolean),
        dailyDate: '2026-06-05',
        monthRange: '2026-06-01 ~ 2026-06-05',
    },
    daily,
    monthly,
};

fs.writeFileSync(OUTPUT_FILE, JSON.stringify(output, null, 2), 'utf-8');
console.log('\n✅ 数据已写入:', OUTPUT_FILE);
console.log('文件大小:', (fs.statSync(OUTPUT_FILE).size / 1024).toFixed(1), 'KB');
