import { SlideShowBG } from "./exports.js";

(function () {
  "use strict"; //strict js to jelp reduce accidental errors like undeclared variables

  var $body = document.querySelector("body"); //selects body tag from html

  // Play initial animations on page load.
  window.addEventListener("load", function () {
    window.setTimeout(function () {
      $body.classList.remove("is-preload");
    }, 100);
  });
  // Slideshow Background.
  SlideShowBG(1);

  //Make the other textbox appear and disappear
  document.getElementById("other").addEventListener("change", function () {
    document.getElementById("other-reason").style.display = this.checked
      ? "block"
      : "none";
  });

  // Calculate the number of working days leave requested
  document
    .getElementById("start-date")
    .addEventListener("change", calculateWorkingDays);
  document
    .getElementById("end-date")
    .addEventListener("change", calculateWorkingDays);

  function calculateWorkingDays() {
    const startDate = new Date(document.getElementById("start-date").value);
    const endDate = new Date(document.getElementById("end-date").value);
    let workingDays = 0;

    if (isNaN(startDate) || isNaN(endDate) || startDate > endDate) {
      document.getElementById("total-days").textContent = "Invalid date range";
      return;
    }

    let currentDate = new Date(startDate);
    while (currentDate <= endDate) {
      const dayOfWeek = currentDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        workingDays++;
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }
    document.getElementById(
      "total-days"
    ).textContent = `Total working days requested: ${workingDays}`;
  }

  document.addEventListener("DOMContentLoaded", () => {
    document
      .querySelector("leave-submit input")
      .addEventListener("click", submitForm());
  });
})();

// Save form data and open email to send formatted document
async function submitForm(event) {
  event.preventDefault(); // Prevent from submission and page reload

  // Get form values
  const name = document.getElementById("name").value;
  const supervisor = document.getElementById("supervisor").value;
  const date = document.getElementById("date").value;
  const startDate = document.getElementById("start-date").value;
  const endDate = document.getElementById("end-date").value;
  const notes = document.getElementById("additional-notes").value;

  const leaveTypes = Array.from(
    document.querySelectorAll('input[name="leaveType"]:checked')
  ).map((box) => box.value);

  const otherReason = document.getElementById("other-reason").value;
  if (leaveTypes.includes("other") && otherReason) {
    leaveTypes.push(`Other: ${otherReason}`);
  }
}

async function getExcelDataApplications(accessToken, siteId, fileId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('Data')/range('')`;

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
