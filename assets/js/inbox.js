import {
  getExcelData,
  SlideShowBG,
  excelSerialDateToJSDate,
} from "./exports.js";

(function () {
  ("use strict"); //strict js to help reduce accidental errors like undeclared variables

  var $body = document.querySelector("body"); //selects body tag from html

  // Play initial animations on page load.
  window.addEventListener("load", function () {
    window.setTimeout(function () {
      $body.classList.remove("is-preload");
    }, 100);
  });
  // Slideshow Background.
  SlideShowBG(1);

  document.addEventListener("DOMContentLoaded", async () => {
    const accessToken = sessionStorage.getItem("accessToken");
    const siteId =
      "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";

    const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
    try {
      const excelData = await getExcelData(accessToken, siteId, fileId);
      const requestList = document.getElementById("request-list");

      if (excelData && requestList) {
        loadInbox(excelData, requestList);
      } else {
        console.error(
          "Error loading Excel data or request list element not found."
        );
      }
    } catch (error) {
      console.error("Error fetching Excel Data:", error);
    }
  });
})();

function loadInbox(excelData, requestList) {
  const requests = [];

  for (let i = 1; i < excelData.length; i++) {
    let name = excelData[i][4]; //Col E
    let supervisor = excelData[i][5]; //Col F
    let date = excelSerialDateToJSDate(excelData[i][6]); //Col G
    let reason = excelData[i][7]; //Col H
    let other = excelData[i][8]; //Col I
    let startDate = excelSerialDateToJSDate(excelData[i][9]); //Col J
    let endDate = excelSerialDateToJSDate(excelData[i][10]); //Col K
    let totalDays = excelData[i][11]; //Col L
    let comments = excelData[i][12]; //Col M
    let approvalStatus = excelData[i][13]; //Col N

    if (!name || !startDate) continue;

    requests.push({
      i,
      name,
      supervisor,
      date,
      reason,
      other,
      startDate,
      endDate,
      totalDays,
      comments,
      approvalStatus,
    });
  }

  requests.forEach((req) => {
    const item = document.createElement("div");
    item.className = "request-item";
    item.innerHTML = `
    <strong>${req.name}</strong><br/>
    ${req.startDate.toDateString()} to ${req.endDate.toDateString()}<br/>
    Status: <em>${req.approvalStatus}</em><br/>
    `;
    requestList.appendChild(item);
  });
}
