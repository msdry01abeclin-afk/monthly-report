const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const yaml = require('js-yaml');
const readlineSync = require('readline-sync');
const os = require('os');

const un = os.userInfo().username;
const yearMonth = readlineSync.question('対象年月（yyyymm）：');
const year = parseInt(yearMonth.slice(0, 4));
const month = parseInt(yearMonth.slice(4, 6));
const fiscalYear = month <= 3 ? year - 1 : year;

const dirpath = String.raw`C:\Users\\${un}\Box\02_PPP事業部\03_PPP_管理G\B_損益\\${fiscalYear}年度\c_実績管理\01_運営部用\03_関西\\${yearMonth}`;
const outputTo = String.raw`C:\Users\\${un}\Box\01_${un}\関西運営室月次まとめ\monthly_yaml`;

const entries = fs.readdirSync(dirpath, {withFileTypes: true});
const files = entries.filter(e => e.isFile() && e.name.endsWith('.xlsm')).map(e => e.name);

const sheetName = '実績_エクセル内計算';
//const commonHeader = ['年度計画', '7月見直し', '10月見直し', '1月見直し', '推定実績', '最新実績'];

const waste_categories = {'受入量':9, '処理量':10}
const total_categories = {'売上':12, '原価':13, '粗利':14};
const cost_item_names = {
    '用役・薬剤': {
        '用役': {
            '電力（買電）': 17,
            'ガス（LNG）': 18,
            'コークス': 19,
            '灯油': 20,
            '上下水道': 21
        },
        '薬剤': {
            '重金属捕集剤': 23,
            '消石灰': 24,
            '苛性ソーダ': 25,
            'アンモニア・尿素水': 26,
            '活性炭': 27
        },
        '灰溶融炉資材': {
            '灰溶融炉資材': 29
        },
        'その他用役・薬剤・消耗品': {
            'その他用役・薬剤・消耗品': 30
        }
    },
    '副産物': {
        '売電': {
            '売電': 32
        },
        '灰処理費用': {
            '灰処理費用': 33
        },
        'その他副産物': {
            'その他副産物':34
        }
    },
    '運転委託': {
        '運転委託': {
            '運転委託':36
        }
    },
    '維持管理': {
        '維持管理': {
            '維持管理': 37
        }
    },
    '工費': {
        '工費':{
            '工費':38
        }
    },
    '経費（予備費含）': {
        '経費（予備費含）': {
            '経費（予備費含）':39
        }
    }
};

const getSiteData = (siteFilepath) => {
    const wb = XLSX.readFile(path.join(dirpath, siteFilepath));
    const ws = wb.Sheets[sheetName];

    return {
        '工場名' :siteFilepath,
        '最新見直し' :'1月見直し',
        '廃棄物' :getSiteWasteTotal(ws, waste_categories),
        '集計' :getSiteWasteTotal(ws, total_categories),
        '費用分類': getSiteCost(ws)
    }
}

const getSiteWasteTotal = (ws, categories) => {
    return Object.entries(categories).map(([cat, r]) =>
        getWasteTotal(ws, cat, r)
    );
}

const getSiteCost = (ws) => {
    return Object.entries(cost_item_names).flatMap(([cat, v1]) =>
        Object.entries(v1).flatMap(([scat, v2]) =>
            Object.entries(v2).map(([itm, r]) =>
                getCost(ws, cat, scat, itm, r)
            )
        )
    );
}

const getWasteTotal = (ws, category, row) => {
    return {    
        '分類' : category,
        'データ': {
            '月次'      : getItems(ws, row, ['H', 'L', 'M', 'N', 'O']),
            '累計'      : getItems(ws, row, ['Q', 'U', 'V', 'W', 'Y'] ),
            '年間見通し'  : getItems(ws, row, ['AA', 'AE', 'AF', 'AG', 'AH']),
        }
    }
}

const getCost = (ws, category, subCategory, item, row) => {
    return {    
        '分類' : category,
        '小分類' : subCategory,
        '項目' : item,
        'データ': {
            '月次'      : getItems(ws, row, ['H', 'L', 'M', 'N', 'O']),
            '累計'      : getItems(ws, row, ['Q', 'U', 'V', 'W', 'Y'] ),
            '年間見通し'  : getItems(ws, row, ['AA', 'AE', 'AF', 'AG', 'AH']),
        }
    }
}

const getItems = (ws, row, targetColumns) => {
    const tvalue = targetColumns.map(v => ws[`${v}${row}`]?.v ?? null);
    return {
        '年度計画' : Math.round(tvalue[0] * 1000)/1000,
        '最新見直し' : Math.round(tvalue[1] * 1000)/1000,
        '実績' : Math.round(tvalue[2] * 1000)/1000,
        '差異' : Math.round(tvalue[3] * 1000)/1000,
        '差異理由・備考' : tvalue[4]
    };
};

const siteData = files.map(f => getSiteData(f))
const dataset = yaml.dump({'年月': yearMonth, '工場データ': siteData});
const BOM = '\uFEFF';
fs.writeFileSync(path.join(outputTo, `${yearMonth}.yml`),
    BOM + dataset,
    'utf-8'
);