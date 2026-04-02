#target illustrator

function exportTextAndCoordinates() {
    if (app.documents.length === 0) {
        alert("請先開啟一個 Illustrator 文件。");
        return;
    }

    var doc = app.activeDocument;
    var outputData = [];
    var textCount = 0;

    // 遞迴函數：用來找出所有物件（包含群組內的）
    function searchItems(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];

            // 如果是文字物件
            if (item.typename === "TextFrame") {
                var content = item.contents.replace(/\r|\n/g, " "); // 取代換行符號避免格式跑掉
                var bounds = item.geometricBounds; // [左, 上, 右, 下]
                
                var info = (textCount + 1) + ". [TextFrame] 內容: \"" + content + "\" - 座標: (" + 
                           bounds[0].toFixed(4) + ", " + 
                           bounds[1].toFixed(4) + ", " + 
                           bounds[2].toFixed(4) + ", " + 
                           bounds[3].toFixed(4) + ")";
                outputData.push(info);
                textCount++;
            } 
            // 如果是群組，進入群組繼續找
            else if (item.typename === "GroupItem") {
                searchItems(item.pageItems);
            }
            // 如果是一般路徑（PathItem），若你需要座標也可以一併紀錄
            else if (item.typename === "PathItem") {
                var pBounds = item.geometricBounds;
                outputData.push((textCount + 1) + ". [PathItem] 未命名物件 - 座標: (" + 
                                pBounds[0].toFixed(4) + ", " + 
                                pBounds[1].toFixed(4) + ", " + 
                                pBounds[2].toFixed(4) + ", " + 
                                pBounds[3].toFixed(4) + ")");
                textCount++;
            }
        }
    }

    // 開始掃描所有頁面物件
    searchItems(doc.pageItems);

    if (outputData.length === 0) {
        alert("找不到任何文字或物件。");
        return;
    }

    // 儲存檔案
    var saveFile = File.saveDialog("儲存匯出的資料", "Text Files: *.txt");
    if (saveFile) {
        saveFile.open("w");
        saveFile.encoding = "UTF-8";
        saveFile.write(outputData.join("\n"));
        saveFile.close();
        alert("匯出完成！共處理 " + outputData.length + " 個物件。");
    }
}

exportTextAndCoordinates();