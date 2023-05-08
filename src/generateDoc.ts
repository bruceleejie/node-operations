
const fs = require('fs');
const officegen = require('officegen');
const docx = officegen('docx');

docx.on('finalize', (written) => {
    console.log(
        "Finish to create Word file.\nTotal bytes created: " + written + "\n"
    );
})

docx.on("error", function (err) {
    console.log(err);
});

var table = [
    [
        {
            val: "No.", opts: { align: "center", vAlign: "center", sz: "20",
            // cellColWidth: 42,
            // b:true,
            // sz: '48',
            // shd: {
            //   fill: "7F7F7F",
            //   themeFill: "text1",
            //   "themeFillTint": "80"
            // },
            // fontFamily: "Avenir Book"
            },
        }, {
            val: "省份", opts: { align: "center", vAlign: "center", sz: "20",
            // b:true,
            // color: "A00000",
            // align: "right",
            // shd: {
            //   fill: "92CDDC",
            //   themeFill: "text1",
            //   "themeFillTint": "80"
            // }
            },
        }, {
            val: "市", opts: { align: "center", vAlign: "center", sz: "20",
            // cellColWidth: 42,
            // b:true,
            // sz: '48',
            // shd: {
            //   fill: "92CDDC",
            //   themeFill: "text1",
            //   "themeFillTint": "80"
            // }
            },
        }, {
            val: "区/县", opts: { align: "center", vAlign: "center", sz: "20",
            // cellColWidth: 42,
            // b:true,
            // sz: '48',
            // shd: {
            //   fill: "92CDDC",
            //   themeFill: "text1",
            //   "themeFillTint": "80"
            //  }
            },
        },
    ],
];
var data = [
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "北京", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "北京", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "上海", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "深圳", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "哈尔滨", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "哈尔滨", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "哈尔滨", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "北京", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "合肥", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "合肥", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "安庆", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "芜湖", cityEn: "beijing",
    },
    {
        id: "101010100", provinceZh: "北京", leaderZh: "北京", cityZh: "芜湖", cityEn: "beijing",
    },
];
var tableStyle = {
    tableColWidth: 2400,
    tableSize: 24,
    tableColor: "ada",
    tableAlign: "center",
    tableVAlign: "center",
    tableFontFamily: "Comic Sans MS",
    borders: true
}

//var tows = ['id', 'provinceZh', 'leaderZh', 'cityZh', 'cityEn'];//创建一个和表头对应且名称与数据库字段对应数据，便于循环取出数据
var pObj = docx.createP({ align: "center" }); // 创建行 设置居中 大标题
pObj.addText("全国所有城市", { bold: true, font_face: "Arial", font_size: 18 }); // 添加文字 设置字体样式 加粗 大小

// let towsLen = tows.length
let dataLen = data.length;
for (var i = 0; i < dataLen; i++) {
  //循环数据库得到的数据，因为取出的数据格式为

  /************************* 文本 *******************************/
//   var pObj = docx.createP();//创建一行，可以
//   pObj.addText(`(${i+1}), `,{ bold: true, font_face: 'Arial',});
//   pObj.addText(`省级:`,{ bold: true, font_face: 'Arial',});
//   pObj.addText(`${data[i]['provinceZh']}  `,);
//   pObj.addText(`市级：`,{ bold: true, font_face: 'Arial',});
//   pObj.addText(`${data[i]['leaderZh']}  `);
//   pObj.addText(`县区：`,{ bold: true, font_face: 'Arial',});
//   pObj.addText(`${data[i]['cityZh']}`);

  /************************* 表格 *******************************/
  let SingleRow = [
    data[i]["id"],
    data[i]["provinceZh"],
    data[i]["leaderZh"],
    {
        val: data[i]["cityZh"],
        opts: {vMerge: i == 0 ? 'restart': (data[i]["cityZh"] != data[i-1]["cityZh"] ? 'restart' :'contiune')}
    },
  ];
  table.push(SingleRow as any);
}
docx.createTable(table, tableStyle);

// word分页
docx.putPageBreak()

// 下一页内容
docx.createTable(table, tableStyle);
var out = fs.createWriteStream("out05051633.docx"); // 文件写入
out.on("error", function (err) {
  console.log(err);
});
var result = docx.generate(out); // 当前目录生成word

