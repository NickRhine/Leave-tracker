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

  // Leave history list
  document.addEventListener("DOMContentLoaded", function () {
    const leaveHistoryList = document.getElementById("leave-history");

    // Example leave data (replace this with real data from storage/API)
    const leaveRecords = [
      "2024-03-01 - 2024-03-03",
      "2024-02-15 - 2024-02-18",
      "2024-01-10 - 2024-01-12",
    ];

    leaveRecords.forEach((leave) => {
      let li = document.createElement("li");
      li.textContent = leave;
      leaveHistoryList.appendChild(li);
    });
  });
})();
