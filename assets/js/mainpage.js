import {
  SlideShowBG,
  getExcelData,
  excelSerialDateToJSDate,
} from "./exports.js";

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

  // document.addEventListener("DOMContentLoaded", function () {
  //   if (window.location.pathname.includes("mainpage.html")) {
  //     document.body.classList.add("no-scroll");
  //   }
  // });

  document.addEventListener("DOMContentLoaded", () => {
    document
      .querySelector(".logout-button input")
      .addEventListener("click", logout);
  });

  // Call this function on page load
  document.addEventListener("DOMContentLoaded", updateLeaveInfo);

  updateNotificationBadge(true); // Shows badge
})();

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

  // const excelData = await getExcelDataLeaveBalance(accessToken, siteId, fileId);
  const excelData = await getExcelData(accessToken, siteId, fileId);

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
      let startDateSerial = excelData[i][9]; // Column I (Index 9) - Start Date
      let endDateSerial = excelData[i][10]; // Column J (Index 10) - End Date
      if (!isNaN(startDateSerial) && !isNaN(endDateSerial)) {
        let startDate = excelSerialDateToJSDate(parseInt(startDateSerial, 10));
        let endDate = excelSerialDateToJSDate(parseInt(endDateSerial, 10));
        let approvalStatus = excelData[i][13]; // Column N (Index 13) - Approval Status
        if (endDate >= today) {
          leaveDates.push({
            start: startDate,
            end: endDate,
            status: approvalStatus,
          });
        }
      }
    }
  }

  if (leaveDates.length > 0) {
    // Sort leave dates by the earliest start date
    leaveDates.sort((a, b) => a.start - b.start);
    let latestLeave = leaveDates[0]; // earliest upcoming leave
    upcomingLeave = `${latestLeave.start.toDateString()} - ${latestLeave.end.toDateString()}`;
  } else {
    leaveDates.status = "No upcoming leave";
  }
  document.querySelector(
    "#app-status"
  ).textContent = `Approval Status: ${leaveDates.status}`;
  document.querySelector("#leave-balance").textContent = leaveBalance;
  document.querySelector("#upcoming-leave").textContent = upcomingLeave;
}

// Logout user and clear session storage
function logout() {
  sessionStorage.clear(); // Clears all session data
  localStorage.clear(); // Clears all local storage data
  window.location.href = "../index.html"; // Redirect to the login page
}

function updateNotificationBadge(hasNewNotifications) {
  const badge = document.getElementById("notificationBadge");
  badge.style.display = hasNewNotifications ? "inline" : "none";
}
