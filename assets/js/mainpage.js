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

  // Call this function on page load
  document.addEventListener("DOMContentLoaded", updateLeaveInfo);
})();

async function getExcelDataLeaveBalance(accessToken, siteId, fileId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('Data')/usedRange`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const data = await response.json();
  if (response.ok) {
    return data.values; //2D array with all of A and B columns
  } else {
    console.error("Error fetching Excel Data:", data);
    return null;
  }
}

async function updateLeaveInfo() {
  const accessToken = sessionStorage.getItem("accessToken");
  const userProfile = sessionStorage.getItem("userProfile");
  const userData = JSON.parse(userProfile);
  const userName = userData.displayName;

  if (!accessToken) {
    console.error("Access Token not found.");
    return;
  }

  const siteId =
    "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";

  const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";

  const excelData = await getExcelDataLeaveBalance(accessToken, siteId, fileId);
  // const excelData = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
  if (!excelData) return;

  let leaveBalance = "Not Found"; // Default value if not found
  let upcomingLeave = "None";
  let leaveDates = [];
  let today = new Date(); // Get today's date

  for (let i = 1; i < excelData.length; i++) {
    if (excelData[i][0] === userName) {
      leaveBalance = excelData[i][1];
      break;
    }
  }

  for (let i = 1; i < excelData.length; i++) {
    if (excelData[i][4] === userName) {
      // Column E (Index 4) - Employee Name
      let startDateSerial = excelData[i][5]; // Column F (Index 5) - Start Date
      let endDateSerial = excelData[i][6]; // Column G (Index 6) - End Date
      if (!isNaN(startDateSerial) && !isNaN(endDateSerial)) {
        let startDate = excelSerialDateToJSDate(parseInt(startDateSerial, 10));
        let endDate = excelSerialDateToJSDate(parseInt(endDateSerial, 10));

        if (endDate >= today) {
          leaveDates.push({ start: startDate, end: endDate });
        }
      }
    }
  }

  if (leaveDates.length > 0) {
    // Sort leave dates by the latest start date
    leaveDates.sort((a, b) => b.start - a.start);
    let latestLeave = leaveDates[0]; // Latest upcoming leave
    upcomingLeave = `${latestLeave.start.toDateString()} - ${latestLeave.end.toDateString()}`;
  }

  document.querySelector("#leave-balance").textContent = leaveBalance;
  document.querySelector("#upcoming-leave").textContent = upcomingLeave;
}

function excelSerialDateToJSDate(serial) {
  const excelEpoch = new Date(1899, 11, 30); // Excel's base date is 1899-12-30
  return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}
