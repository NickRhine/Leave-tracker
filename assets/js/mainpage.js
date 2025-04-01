import { SlideShowBG } from "./exports.js";

(function () {
  "use strict"; //strict js to jelp reduce accidental errors like undeclared variables

  var $body = document.querySelector("body"); //selects body tag from html

  //Get user data
  document.addEventListener("DOMContentLoaded", function () {
    // Get user data from sessionStorage
    const userProfile = sessionStorage.getItem("userProfile");

    if (userProfile) {
      const userData = JSON.parse(userProfile);
      document.querySelector(
        "#user-name"
      ).textContent = `Hello, ${userData.displayName}`;
    } else {
      console.error("User data not found in sessionStorage.");
    }
  });

  // Play initial animations on page load.
  window.addEventListener("load", function () {
    window.setTimeout(function () {
      $body.classList.remove("is-preload");
    }, 100);
  });
  // Slideshow Background.
  SlideShowBG(1);

  // Leave history list
  document.addEventListener("DOMContentLoaded", function () {
    const upcomingLeave = document.getElementById("upcoming-leave");

    // Example leave data (replace this with real data from storage/API)
    const leaveRecords = [
      "2024-03-01 - 2024-03-03",
      "2024-02-15 - 2024-02-18",
      "2024-01-10 - 2024-01-12",
    ];

    if (leaveRecords.length > 0) {
      upcomingLeave.textContent = leaveRecords[0];
    } else {
      upcomingLeave.textContent = "No upcoming leave";
    }
  });

  // Call this function on page load
  document.addEventListener("DOMContentLoaded", updateLeaveInfo);
})();

async function getSiteId(accessToken) {
  const url = `https://graph.microsoft.com/v1.0/sites/netorg7968809.sharepoint.com:/sites/RHINEmechatronics-BusinessDevelopment`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const data = await response.json();
  if (response.ok) {
    return data.id;
  } else {
    console.error("Error fetching Site ID:", data);
    return null;
  }
}

async function getFileId(accessToken, siteId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/Rename.xlsx`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const data = await response.json();
  if (response.ok) {
    return data.id;
  } else {
    console.error("Error fetching File ID:", data);
    return null;
  }
}

async function getExcelData(accessToken, siteId, fileId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('Data')/range(address='A1:A2')`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const data = await response.json();
  if (response.ok) {
    return data.values;
  } else {
    console.error("Error fetching Excel Data:", data);
    return null;
  }
}

async function updateLeaveInfo() {
  const accessToken = sessionStorage.getItem("accessToken");

  if (!accessToken) {
    console.error("Access Token not found.");
    return;
  }

  // const siteId = await getSiteId(accessToken);
  // if (!siteId) return;
  const siteId =
    "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";

  const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
  // "b!lFDv1l-H10eTxEOuFxoE_yGBOoh0A_RJlHYtO5ocs4oWLWqGN5CiRLBtt1hFEdjV";
  // if (!fileId) return;

  const excelData = await getExcelData(accessToken, siteId, fileId);
  // const excelData = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
  // if (!excelData) return;

  // Assuming column A is "Leave Type" and column B is "Days Remaining"
  document.querySelector("#leave-balance").textContent = excelData[0][0]; // Adjust index based on your Excel structure
  document.querySelector("#upcoming-leave").textContent = excelData[1][0]; // Adjust index based on your Excel structure
}
