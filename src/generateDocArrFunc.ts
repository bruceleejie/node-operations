
const fs = require('fs');
const officegen = require('officegen');
const docx = officegen('docx');

// let fileJson = fs.readFileSync(__dirname+'\\file\\source.json', 'utf-8');
let fileJson = fs.readFileSync(__dirname+'\\file\\source.json', 'utf-8');
// console.log(6, fileJson);
let arr = JSON.parse(fileJson);

docx.on('finalize', (written) => {
    console.log(
        "Finish to create Word file.\nTotal bytes created: " + written + "\n"
    );
})

docx.on("error", function (err) {
    console.log(err);
});

arr.forEach(element => {
    let pObj = docx.createP({ align: "center" }); // 创建行 设置居中 大标题
    pObj.addText(element.title, { bold: true, font_face: "Arial", font_size: 18 }); // 添加文字 设置字体样式 加粗 大小
    element.children.forEach(childitem => {
        let tableHeadArr = childitem.config['表格字段归一化结果']['表头顺序'].map(item => {
            let obj = {
                val: item.title, opts: { align: "center", vAlign: "center", sz: "20",},
            }
            return obj;
        })
        let tableStyle = {
            tableColWidth: 2400,
            tableSize: 24,
            tableColor: "ada",
            tableAlign: "left",
            tableVAlign: "center",
            tableFontFamily: "Comic Sans MS",
            borders: true
        }
        let table:any = [];
        table[0] = tableHeadArr;
        let dataArr:any = [];
        childitem.config['对比结果']&&childitem.config['对比结果'].forEach(resultitem => {
            dataArr.push(resultitem);
        });
        let pObj = docx.createP({ align: "left" }); // 创建行 设置居中 大标题
        pObj.addText(childitem.title, { bold: true, font_face: "Arial", font_size: 16 }); // 添加文字 设置字体样式 加粗 大小
        let dataLen = dataArr.length;
        for (var i = 0; i < dataLen; i++) {
            let bigTrItem = dataArr[i]; // 一个大行，children里有多个小行
            bigTrItem.children.forEach((smallTrItem,index)=>{
                let SingleRow : any = [];
                tableHeadArr.forEach(item => {
                    if(index==0) {
                        SingleRow.push({
                            val: smallTrItem[item.val],
                            opts: {vMerge: 'restart'}
                        })
                    } else {
                        if(smallTrItem[item.val]==bigTrItem[item.val]) {
                            SingleRow.push({
                                val: smallTrItem[item.val],
                                opts: {vMerge: 'continue'}
                            })
                        }else{
                            if(smallTrItem[item.val]==bigTrItem.children[index-1][item.val]) {
                                SingleRow.push({
                                    val: smallTrItem[item.val],
                                    opts: {vMerge: 'continue'}
                                })
                            }else {
                                SingleRow.push({
                                    val: smallTrItem[item.val],
                                    opts: {vMerge: 'restart'}
                                })
                            }
                        }
                    }
                })
                table.push(SingleRow as any);
            })
        }
        docx.createTable(table, tableStyle);
        // word分页
        docx.putPageBreak()
    });
});

var out = fs.createWriteStream(`result_${new Date().getTime()}.docx`); // 文件写入
out.on("error", function (err) {
  console.log(err);
});
var result = docx.generate(out); // 当前目录生成word


