const fs = require('fs');
const path = require('path');
const readlineSync = require('readline-sync');

const yearMonth = readlineSync.question('対象年月（yyyymm）：');

const dirpath = String.raw`C:\Users\jfe106005\Box\02_PPP事業部\03_PPP_管理G\B_損益\2025年度\c_実績管理\01_運営部用\03_関西` + `\\${yearMonth}`;
const entries = fs.readdirSync(dirpath, {withFileTypes: true});

const files = entries.filter(e => e.isFile()).map(e => e.name);

const sheetName = '実績_エクセル内計算';
const commonHeader = ['年度計画', '7月見直し', '10月見直し', '1月見直し', '推定実績', '最新実績'];

const dataset = [];

files.forEach(f => {
    dataset.push(getSiteData(path.join(dirpath, f), f, sheetName, commonHeader));
});

//console.log(dataset);

const BOM = '\uFEFF';
fs.writeFileSync(dirpath +`\\output.json`,
     BOM + JSON.stringify(dataset, null, 2),
     'utf-8'
    );

function getSiteData(pathname, filename, sheetName, commonHeader)
{
    const XLSX = require('xlsx');
    const wb = XLSX.readFile(pathname);
    const ws = wb.Sheets[sheetName];

    function getSheetData(sheet, range) {
        return XLSX.utils.sheet_to_json(sheet, { header: commonHeader, range })
        .map(row =>
        Object.fromEntries(
            Object.entries(row).map(([key, val]) => [
            key,
            typeof val === 'number' ? Math.round(val) : val
            ])
        )
        );
    }

    const Monthly = getSheetData(ws, "H14:M14")[0];
    const YTD = getSheetData(ws, "Q14:V14")[0];
    const Landing = getSheetData(ws, "AA14:AF14")[0];

    const data = {
        '当月損益': Monthly,
        '累計損益': YTD,
        '年間見通し': Landing
    }

    return {'ファイル':filename, '粗利':data};
}