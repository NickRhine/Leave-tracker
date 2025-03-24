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
})();
