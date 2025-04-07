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
}

//Generate a pdf using a screenshot of the form
function generatePDF() {
  const userProfile = sessionStorage.getItem("userProfile");
  let username = ""; // Default value
  if (userProfile) {
    const userData = JSON.parse(userProfile);
    username = userData.displayName;
  } else {
    console.error("User data not found in sessionStorage.");
  }

  const element = document.body; // The content you want to convert into PDF
  const opt = {
    margin: 1,
    filename: `leave_request_${username}.pdf`,
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
  };

  // Generate PDF
  html2pdf().from(element).set(opt).save();
}

// Attach pdf to email and open email client
function sendEmailWithPDF(pdfData) {
  emailjs
    .send("YOUR_SERVICE_ID", "YOUR_TEMPLATE_ID", {
      user_email: "nicholas.hobden@rhinemechatronics.com", // The email of the recipient
      subject: "Leave Request",
      message: "Please find attached the leave request PDF.",
      attachment: pdfData, // Send the PDF as an attachment
    })
    .then((response) => {
      console.log("Email sent successfully:", response);
      alert("Email sent successfully!");
    })
    .catch((error) => {
      console.error("Error sending email:", error);
    });
}
