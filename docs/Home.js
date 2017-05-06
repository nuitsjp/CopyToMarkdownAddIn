/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // FabricUI 通知メカニズムを初期化して、非表示にします
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Excel 2016 を使用していない場合は、フォールバック ロジックを使用してください。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("このサンプルでは、スプレッドシートで選ばれたセルの値が表示されます。");
                $('#button-text').text("表示!");
                $('#button-desc').text("選択範囲が表示されます");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("このサンプルでは、スプレッドシートで選択したセルから最も高い値が強調表示されます。");
            $('#button-text').text("強調表示!");
            $('#button-desc').text("最大数が強調表示されます。");
                
            loadSampleData();

            // 強調表示ボタンのクリック イベント ハンドラーを追加します。
            $('#highlight-button').click(hightlightHighestValue);
        });
    }

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Excel オブジェクト モデルに対してバッチ操作を実行します
        Excel.run(function (ctx) {
            // 作業中のシートに対するプロキシ オブジェクトを作成します
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // ワークシートにサンプル データを書き込むコマンドをキューに入れます
            sheet.getRange("B3:D5").values = values;

            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Excel オブジェクト モデルに対してバッチ操作を実行します
        Excel.run(function (ctx) {
            // 選択された範囲に対するプロキシ オブジェクトを作成し、そのプロパティを読み込みます
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // セルを検索して強調表示します
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // セルを強調表示
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('選択されたテキスト:', '"' + result.value + '"');
                } else {
                    showNotification('エラー', result.error.message);
                }
            });
    }

    // エラーを処理するためのヘルパー関数
    function errorHandler(error) {
        // Excel.run の実行から浮かび上がってくるすべての累積エラーをキャッチする必要があります
        showNotification("エラー", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 通知を表示するヘルパー関数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
