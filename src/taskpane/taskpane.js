/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;
    document.getElementById("getMeetingLinkButton").onclick = ssoAuthHelper.getMeeting;
    document.getElementById("getCellSumButton").onclick = writeSumToDocument;
  }
});

function writeSumToDocument() {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeSource = sheet.getRange("A1:B1");
    rangeSource.load("values");
    return context.sync().then(() => {
      console.log(rangeSource.values);
      var sum = rangeSource.values[0][0] + rangeSource.values[0][1];
      const rangeSumAddress = "A2:B2";
      const range = sheet.getRange(rangeSumAddress);
      range.values = [["sum",sum]];
      range.format.autofitColumns();
    });
  });
}
