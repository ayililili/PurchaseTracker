/* global Excel, console */

export async function testWorksheetChange() {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const sheet = context.workbook.worksheets.getItem("Sheet1");

      // 設置工作表變更事件監聽器，並使用 async 以符合 Promise 型別
      sheet.onChanged.add(async (args: Excel.WorksheetChangedEventArgs) => {
        console.log("Worksheet changed!");

        // 可以檢查變更的儲存格範圍
        console.log("Changed address:", args.address);
      });

      await context.sync();
    });
  } catch (error) {
    console.error("Error setting up worksheet change monitor:", error);
  }
}
