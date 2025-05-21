import {
  getExcelData,
  SlideShowBG,
  excelSerialDateToJSDate,
  updateExcelRow,
} from "./exports.js";

const accessToken = sessionStorage.getItem("accessToken");
const siteId =
  "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";

const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";

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

// Takes excel data and populates the inbox with requests when item in inbox is clicked it expands to show more details
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

    if (!name || !startDate || startDate < new Date()) continue; //Skip empty rows and dates which have passed

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
  if (requests.length === 0) {
    const noRequests = document.createElement("div");
    noRequests.className = "no-requests";
    noRequests.innerHTML = "<p>No requests at the moment.</p>";
    requestList.appendChild(noRequests);
    return;
  }

  requests.forEach((req) => {
    const item = document.createElement("div");
    item.className = "request-item";

    item.innerHTML = `
    <strong>${req.name}</strong><br/>
    ${req.startDate.toDateString()} to ${req.endDate.toDateString()}<br/>
    Status: <em>${req.approvalStatus}</em><br/>
    `;

    const details = document.createElement("div");
    details.className = "request-details";
    details.style.display = "none";

    details.innerHTML = `
    <p><strong>Supervisor:</strong> ${req.supervisor}</p>
  <p><strong>Requested on:</strong> ${req.date.toDateString()}</p>
  <p><strong>Reason:</strong> ${req.reason}</p>
  <p><strong>If other:</strong> ${req.other}</p>
  <p><strong>Total Days:</strong> ${req.totalDays}</p>
  <p><strong>Comments:</strong> ${req.comments || "None"}</p>
  <div class="actions">
    <button class="approve-btn">Approve</button>
    <button class="deny-btn">Deny</button>
  </div>
  `;

    item.appendChild(details);
    item.addEventListener("click", () => {
      const isVisible = details.style.display === "block";
      details.style.display = isVisible ? "none" : "block";
    });

    const rowData = [
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ];
    details.querySelector(".approve-btn").addEventListener("click", () => {
      rowData[9] = "Approved";
      // Call the function to update the Excel sheet with the new status and comments
      updateExcelRow(accessToken, siteId, fileId, req.i, rowData);
    });

    details.querySelector(".deny-btn").addEventListener("click", () => {
      rowData[9] = "Denied";
      // Call the function to update the Excel sheet with the new status and comments
      updateExcelRow(accessToken, siteId, fileId, req.i, rowData);
    });

    requestList.appendChild(item);
  });
}
