// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
(function () {
    Office.initialize = function (reason) {
        // 必要な初期化は、ここで実行できます。
    };
})();

var NewLine = "\n";

function copyToMarkdown(event) {
    Excel.run(function (ctx) {
        var cells = [];
        var range = ctx.workbook.getSelectedRange().load(["rowCount", "columnCount"]);
        return ctx.sync()
            .then(function () {
                for (var row = 0; row < range.rowCount; row++) {
                    for (var col = 0; col < range.columnCount; col++) {
                        cells.push(range.getCell(row, col).load(["text", "format"]));
                    }
                }
            })
            .then(ctx.sync)
            .then(function() {
                var resultBuffer = new StringBuilder();
                var separatorBuffer = new StringBuilder();
                for (var x = 0; x < range.columnCount; x++)
                {
                    var cell = cells[x];

                    resultBuffer.append("|");
                    resultBuffer.append(formatText(cell.text));
                    switch (cell.format.horizontalAlignment)
                    {
                        case "Center":
                            separatorBuffer.append("|:-:");
                            break;
                        case "Right":
                            separatorBuffer.append("|--:");
                            break;
                        default:
                            separatorBuffer.append("|:--");
                            break;
                    }
                }
                resultBuffer.append("|");
                resultBuffer.append(NewLine);
                separatorBuffer.append("|");
                separatorBuffer.append(NewLine);
                resultBuffer.append(separatorBuffer.toString());

                for (var row = 1; row < range.rowCount; row++)
                {
                    for (var col = 0; col < range.columnCount; col++)
                    {
                        var valueCell = cells[row * range.columnCount + col];

                        resultBuffer.append("|");
                        resultBuffer.append(formatText(valueCell.text));
                    }
                    resultBuffer.append("|");
                    resultBuffer.append(NewLine);
                }

                var result = resultBuffer.toString();
                window.clipboardData.setData("Text", result);
            });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    event.completed();
}

function formatText(range)
{
    if (range == undefined) {
        return "";
    }
    else {
        return range.join().replace(NewLine, "<br>");
    }
}


var StringBuilder = function (string) {
    this.buffer = [];

    this.append = function (string) {
        this.buffer.push(string);
        return this;
    };

    this.toString = function () {
        return this.buffer.join('');
    };

    if (string) {
        this.append(string);
    }
};
