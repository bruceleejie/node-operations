
const fs = require('fs');
const officegen = require('officegen');
// const docx = officegen('docx');
const pageStyle = {
    type: 'docx',
    pageMargins: {
        top: 1000,
        bottom: 1000,
        left: 1000,
        right: 1000,
    }
}
const docx = officegen(pageStyle);

exports.generateTableDoc = function (arr = [],outputPath) {
    docx.on('finalize', (written) => {
        console.log(
            "Finish to create Word file.\nTotal bytes created: " + written + "\n"
        );
    })
    
    docx.on("error", function (err) {
        console.log(err);
    });
    
    arr.forEach(element => {
        // let levelOneObj = docx.createNestedOrderedList({
        //     "level": 1, align: "left"
        // })
        let levelOneObj = docx.createP({ align: "left" }); // 创建行 设置居中 大标题
        levelOneObj.addText(element.title, { bold: true, font_face: "楷体", font_size: 18 }); // 添加文字 设置字体样式 加粗 大小
        element.children.forEach(childitem => {
            let tableHeadArr = []
            let tableStyle = {}
            if(childitem.config['展示方式'] != '纵向') {
                tableHeadArr = childitem.config['表格字段归一化结果']['表头顺序'].map(item => {
                    let obj = {
                        val: item.title, opts: { align: "center", vAlign: "center", sz: "22", b: true, shd: {
                            fill: "e9e9e9",
                        }},
                    }
                    return obj;
                })
                tableStyle = {
                    tableColWidth: 2400,
                    tableSize: 22,
                    tableColor: "ada",
                    tableAlign: "left",
                    tableVAlign: "center",
                    tableFontFamily: "楷体",
                    borders: true,
                    columns:[         // 单列样式配置
                        {width: 1000}
                    ]
                }
                tableHeadArr.forEach((headitem, headindex) => {
                    if(headindex == 0) {
                        tableStyle.columns[0] = {width: 1000};
                    } else {
                        tableStyle.columns.push({width: 2400});
                    }
                })
            } else {
                tableHeadArr = childitem.config['表格字段归一化结果']['表头顺序'].map(item => {
                    let obj = {
                        val: item.title, 
                        opts: { align: "center", vAlign: "center", sz: "22", b: true, fontFamily: "楷体", shd: {
                            fill: "e9e9e9",
                        }},
                    }
                    return obj;
                })
                tableStyle = {
                    tableColWidth: 2400,
                    tableSize: 22,
                    tableColor: "ada",
                    tableAlign: "left",
                    tableVAlign: "center",
                    tableFontFamily: "楷体",
                    borders: true,
                }
            }
            
            let table = []; // 横表格
            let columnTable = []; // 纵表格
            if(childitem.config['展示方式'] != '纵向') {
                table[0] = tableHeadArr;
            }
            let dataArr = [];
            childitem.config['对比结果']&&childitem.config['对比结果'].forEach(resultitem => {
                dataArr.push(resultitem);
            });
            // let levelTwoObj = docx.createNestedOrderedList({
            //     "level": 2, align: "left"
            // })
            let levelTwoObj = docx.createP({ align: "left" }); // 创建行 设置居中 大标题
            levelTwoObj.addText(childitem.title, { bold: true, font_face: "楷体", font_size: 16 }); // 添加文字 设置字体样式 加粗 大小
            let dataLen = dataArr.length;
            for (var i = 0; i < dataLen; i++) {
                let bigTrItem = dataArr[i]; // 一个大行，children里有多个小行
                if(childitem.config['展示方式'] == '纵向') {
                    tableHeadArr.forEach((item, index) => {
                        let SingleRow = [];
                        SingleRow.push(item);
                        let firstNullFlag = bigTrItem.children[0][item.val] == undefined || bigTrItem.children[0][item.val] == 'N/A';
                        let secondNullFlag = bigTrItem.children[1][item.val] == undefined || bigTrItem.children[1][item.val] == 'N/A';
                        let thirdNullFlag = bigTrItem.children[2][item.val] == undefined || bigTrItem.children[2][item.val] == 'N/A';
                        if(firstNullFlag) {
                            // 第1项为空
                            SingleRow.push({
                                val: bigTrItem.children[0][item.val] == undefined ? 'N/A' : bigTrItem.children[0][item.val],
                                opts: {align: "center", vAlign: "center", sz: "22",}
                            })
                            if(secondNullFlag || thirdNullFlag) {
                                // 第2 第3项任意一个为空
                                SingleRow.push({
                                    val: bigTrItem.children[1][item.val],
                                    opts: {align: "center", vAlign: "center", sz: "22",}
                                })
                                SingleRow.push({
                                    val: bigTrItem.children[2][item.val],
                                    opts: {align: "center", vAlign: "center", sz: "22",}
                                })
                            } else {
                                // 2 3 项都不为空
                                if(bigTrItem.children[1][item.val] == bigTrItem.children[2][item.val]) {
                                    SingleRow.push({
                                        val: bigTrItem.children[2][item.val],
                                        opts: {align: "center", vAlign: "center", sz: "22", gridSpan: 2}
                                    })
                                } else {
                                    SingleRow.push({
                                        val: bigTrItem.children[1][item.val],
                                        opts: {align: "center", vAlign: "center", sz: "22",}
                                    })
                                    SingleRow.push({
                                        val: bigTrItem.children[2][item.val],
                                        opts: {align: "center", vAlign: "center", sz: "22",}
                                    })
                                }
                            }
                        } else {
                            // 第1项不为空
                            if(secondNullFlag) {
                                // 2项为空
                                SingleRow.push({
                                    val: bigTrItem.children[0][item.val],
                                    opts: {align: "center", vAlign: "center", sz: "22",}
                                })
                                SingleRow.push({
                                    val: bigTrItem.children[1][item.val] == undefined ? 'N/A' : bigTrItem.children[1][item.val],
                                    opts: {align: "center", vAlign: "center", sz: "22",}
                                })
                                SingleRow.push({
                                    val: bigTrItem.children[2][item.val],
                                    opts: {align: "center", vAlign: "center", sz: "22",}
                                })
                            } else {
                                // 2项不为空
                                if(thirdNullFlag) {
                                    // 3项为空
                                    if(bigTrItem.children[0][item.val] == bigTrItem.children[1][item.val]) {
                                        // 1==2列
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", gridSpan: 2}
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[2][item.val] == undefined ? 'N/A' : bigTrItem.children[2][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                    } else {
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[1][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[2][item.val] == undefined ? 'N/A' : bigTrItem.children[2][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                    }
                                } else {
                                    // 3项不为空
                                    if(bigTrItem.children[0][item.val] == bigTrItem.children[1][item.val] && bigTrItem.children[1][item.val] == bigTrItem.children[2][item.val]) {
                                        // 1==2 && 2==3
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", gridSpan: 3}
                                        })
                                    } else if(bigTrItem.children[0][item.val] == bigTrItem.children[1][item.val] && bigTrItem.children[1][item.val] != bigTrItem.children[2][item.val]) {
                                        // 1==2 && 2!=3
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", gridSpan: 2}
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[2][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22",}
                                        })
                                    } else if(bigTrItem.children[0][item.val] != bigTrItem.children[1][item.val] && bigTrItem.children[1][item.val] == bigTrItem.children[2][item.val]) {
                                        // 1!=2 && 2==3
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[1][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", gridSpan: 2}
                                        })
                                    } else {
                                        SingleRow.push({
                                            val: bigTrItem.children[0][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[1][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                        SingleRow.push({
                                            val: bigTrItem.children[2][item.val],
                                            opts: {align: "center", vAlign: "center", sz: "22", }
                                        })
                                    }
                                }
                            }
                        }
                        columnTable.push(SingleRow);
                    })
                } else {
                    bigTrItem.children.forEach((smallTrItem,index)=>{
                        let SingleRow = [];
                        tableHeadArr.forEach((item, headindex) => {
                            if(index==0) {
                                SingleRow.push({
                                    val: smallTrItem[item.val] == undefined ? 'N/A' : smallTrItem[item.val],
                                    opts: {vMerge: 'restart', sz: "22",}
                                })
                            } else {
                                if(smallTrItem[item.val] == 'N/A' || smallTrItem[item.val] == undefined) {
                                    SingleRow.push({
                                        val: smallTrItem[item.val] == undefined ? 'N/A' : smallTrItem[item.val],
                                        opts: {vMerge: 'restart', sz: "22",}
                                    })
                                } else if(smallTrItem[item.val]==bigTrItem[item.val]) {
                                    SingleRow.push({
                                        val: smallTrItem[item.val],
                                        opts: {vMerge: 'continue', sz: "22",}
                                    })
                                }else{
                                    console.log(255, smallTrItem[item.val], Boolean(smallTrItem[item.val]));
                                    let itemDeleteSpace = Boolean(smallTrItem[item.val]) ? (smallTrItem[item.val]).replace(/[()\\\[\]<>（）{}【】\s]/g,"") : smallTrItem[item.val];
                                    let lastItemDeleteSpace = bigTrItem.children[index-1][item.val]&&(bigTrItem.children[index-1][item.val]).replace(/[()\\\[\]<>（）{}【】\s]/g,"");
                                    // if(smallTrItem[item.val]==bigTrItem.children[index-1][item.val]) {
                                    if(itemDeleteSpace == lastItemDeleteSpace) {
                                        SingleRow.push({
                                            val: smallTrItem[item.val],
                                            opts: {vMerge: 'continue', sz: "22",}
                                        })
                                    }else {
                                        SingleRow.push({
                                            val: smallTrItem[item.val],
                                            opts: {vMerge: 'restart', sz: "22",}
                                        })
                                    }
                                }
                            }
                        })
                        table.push(SingleRow);
                    })
                }
            }
            if(table.length > 1) {
                docx.createTable(table, tableStyle);
                // word分页
                docx.putPageBreak()
            } else if(columnTable.length > 0) {
                docx.createTable(columnTable, tableStyle);
                docx.putPageBreak()
            } else {
                let nodataTextObj = docx.createP({ align: "left" }); // 创建行 设置居中 大标题
                nodataTextObj.addText('此部分未找到任何可比对内容。', { bold: false, font_face: "楷体", font_size: 11 }); // 添加文字 设置字体样式 加粗 大小
            }
        });
    });
    
    var out = fs.createWriteStream(outputPath); // 文件写入
    out.on("error", function (err) {
      console.log(err);
    });
    var result = docx.generate(out); // 当前目录生成word
}


