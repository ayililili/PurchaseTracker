/* global Excel, console */

export async function monitorCellAndSync() {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      // 選擇表單 1 並監聽變更事件
      const sourceSheet: Excel.Worksheet = context.workbook.worksheets.getItem("Sheet1");

      // 設置工作表的變更事件監聽器
      sourceSheet.onChanged.add(async (event: Excel.WorksheetChangedEventArgs) => {
        console.log("Change detected in Sheet1");

        try {
          // 確認變更的儲存格包含 A1
          if (event.address === "Sheet1!A1") {
            console.log("Change detected in cell A1");

            // 取得 A1 的範圍並加載其值
            const sourceRange = sourceSheet.getRange("A1");
            sourceRange.load("values"); // 加載變更值
            await context.sync();

            const newValue = sourceRange.values[0][0];

            // 將新值同步到表單 2 的目標儲存格
            const targetSheet: Excel.Worksheet = context.workbook.worksheets.getItem("Sheet2");
            const targetRange = targetSheet.getRange("B1");
            targetRange.values = [[newValue]];
            targetRange.format.autofitColumns();
            await context.sync();
          }
        } catch (error) {
          console.error("Error updating target cell:", error);
        }
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error setting up cell monitor:", error);
  }
}
