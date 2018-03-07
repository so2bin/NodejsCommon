/********************************************************
*   通过高德地图查询到指定地址类型的地址数据，导出到excel表中
*/
const request = require('request');
const qs = require('querystring');
const xlsx = require('node-xlsx').default;
const fs = require('fs');
const XLSXWriter = require('xlsx-writestream');

const URL = 'http://restapi.amap.com/v3/place/text'
const GEOKEY = '';
const oExcelFileName = './data.xlsx'
const OFFSET = 20;    // 每次请求20份数据
const configFile = 'C:\\Users\\Administrator\\Desktop\\data.xlsx';

const TYPES = [
    "120000",   // 商务住宅
    "060000",  // 购物服务
    "060000",  // 餐饮服务
    "100000",  // 住宿服务
].join('|');

let params = {
    keywords: '',
    types: TYPES,
    city: '北京',
    citylimit: true,
    offset: OFFSET,
    page: 1,
    extensions: "all",
    key: GEOKEY
};

function calcQryUrl(params){
    return `${URL}?${qs.stringify(params)}`;
}

function initWriter(writer){
    writer.getReadStream().pipe(fs.createWriteStream(oExcelFileName));
    writer.defineColumns([
        {width: 20},
        {width: 12},
        {width: 12},
        {width: 10},
        {width: 10},
        {width: 10},
        {width: 10},
    ]);
}
function endWriter(writer){
    writer.finalize();
}

function getConfigData(){
    let data = [];
    const workSheetsFromFile = xlsx.parse(configFile);
    const rows = workSheetsFromFile[0].data;
    for(let row of rows){
        data.push({
            region: row[0],
            num: row[1]
        })
    }
    return data;
}

function qryOnce(page, region, writer){
    params.page = page;
    params.city = region;
    return new Promise((resolve, reject)=>{
        request(calcQryUrl(params), (error, response, body)=>{
            if(error){
                reject(error);
            }
            let data = JSON.parse(body);
            if(response.statusCode !== 200){
                reject(`返回错误，${response.statusCode}`)
            }
            let excelRows = [];
            for(let row of data.pois){
                let pos = row.location.split(',');
                if(pos.length != 2){
                    continue;
                }
                writer.addRow({
                    name: row.name,
                    longitude: pos[0],
                    latitude: pos[1],
                    adcode: row.adcode,
                    adname: row.adname,
                    cityname: row.cityname,
                    pname: row.pname
                })
            }
            resolve(data.count - page*OFFSET)
        })
    });
}

async function run(writer, region, maxLeftNum){
    let page = 1;
    let qryLeftNum = await qryOnce(page, region, writer);
    maxLeftNum -= OFFSET;
    let leftNum = Math.min(maxLeftNum, qryLeftNum);
    while(leftNum > 0){
        console.log(`region: ${region}, page: ${page}, left count: ${page*OFFSET}/${leftNum}`)
        page += 1;
        await qryOnce(page, region, writer);
        leftNum -= OFFSET;
        // 限制每秒发送频率
        await new Promise(function(resolve, reject) {
            setTimeout(resolve, 0);
        });
    }
}

async function main(){
    let writer = new XLSXWriter();
    initWriter(writer);
    let confRows = getConfigData();
    for(let row of confRows){
        if(!row || !row.num){
            continue;
        }
        await run(writer, row.region, row.num);
    }
    endWriter(writer)
}

main();
