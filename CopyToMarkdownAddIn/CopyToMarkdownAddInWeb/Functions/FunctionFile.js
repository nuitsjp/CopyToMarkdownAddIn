// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
(function () {
    Office.initialize = function (reason) {
        // 必要な初期化は、ここで実行できます。
    };
})();

function copyToMarkdown(event) {
    Excel.run(function (ctx) {
        var cells = [];
        var range = ctx.workbook.getSelectedRange().load(["rowCount", "columnCount"]);
        return ctx.sync()
            .then(function () {
                for (var row = 0; row < range.rowCount; row++) {
                    for (var col = 0; col < range.columnCount; col++) {
                        cells.push(range.getCell(row, col).load(["address", "text", "values", "format"]));
                    }
                }
            })
            .then(ctx.sync)
            .then(function() {
                var resultBuffer = new StringBuilder();
                var separatorBuffer = new StringBuilder();
                for (var x = 0; x < range.columnCount; x++)
                {
                    var cell = cells[1, x];

                    resultBuffer.append("|");
                    resultBuffer.append(cell.text);
                    console.log("address:" + cell.address);
                    console.log("values:" + cell.values);
                    console.log("text:" + cell.text);
                    var format = cell.format;
                    console.log(format.horizontalAlignment);
                }
                var result = resultBuffer.toString();
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
    if (range == null || range.text == null) {
        return "";
    }
    else {
        return range.text;
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
