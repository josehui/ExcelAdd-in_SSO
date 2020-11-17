/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
/* global $, document, Excel, Office */


import { getGraphData, addCellFunction } from "./../helpers/ssoauthhelper";

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(document).ready(function() {
      $("#getGraphDataButton").click(getGraphData);
      $("#AddCellButton").click(addCellFunction);
    });
  }
});

export function writeDataToOfficeDocument2(): Promise<any> {
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

export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo: string[] = [];
    userProfileInfo.push(result["displayName"]);
    userProfileInfo.push(result["jobTitle"]);
    userProfileInfo.push(result["mail"]);
    userProfileInfo.push(result["mobilePhone"]);
    userProfileInfo.push(result["officeLocation"]);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}