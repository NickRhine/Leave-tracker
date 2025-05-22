import { SlideShowBG, updateExcelRow, getExcelData } from "./exports.js";

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
  // document.getElementById("other").addEventListener("change", function () {
  //   document.getElementById("other-reason").style.display = this.checked
  //     ? "block"
  //     : "none";
  // });

  document.querySelectorAll('input[name="leaveType"]').forEach((radio) => {
    radio.addEventListener("change", function () {
      const otherReason = document.getElementById("other").checked;
      document.getElementById("other-reason").style.display = otherReason
        ? "block"
        : "none";
    });
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
      .querySelector(".leave-submit input")
      .addEventListener("click", submitForm);
  });
})();

// Save form data to excel sheet and send email notification
async function submitForm(event) {
  // #region Submission message declaration
  const $form = document.getElementById("leave-form");
  let $message = $form.querySelector(".message");
  if (!$message) {
    $message = document.createElement("span");
    $message.classList.add("message");
    $form.appendChild($message);
  }

  $message._show = function (type, text) {
    $message.innerHTML = text;
    $message.classList.add(type);
    $message.classList.add("visible");

    setTimeout(function () {
      $message._hide();
    }, 5000);
  };

  $message._hide = function () {
    $message.classList.remove("visible");
  };
  // #endregion

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
    leaveTypes.push(`${otherReason}`);
  }

  // Check if all required fields are filled
  let valid = validateInputs(
    name,
    supervisor,
    date,
    startDate,
    endDate,
    leaveTypes,
    otherReason
  );
  if (!valid) {
    $message._show("failure", "Submission failed");
    return;
  }

  const accessToken = sessionStorage.getItem("accessToken");
  const email = sessionStorage.getItem("email");

  //Collect excel data
  const siteId =
    "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";
  const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
  let data = await getExcelData(accessToken, siteId, fileId, "Data");

  if (data == null) {
    console.error("Error fetching Excel data");
    return;
  }

  //Find empty row
  let emptyRow = null;
  for (let i = 1; i < data.length; i++) {
    //Check employee name column
    if (data[i][0] === "") {
      emptyRow = i;
      break;
    }
  }
  if (emptyRow === null) {
    //If no rows are empty, append to the end of the data
    emptyRow = data.length;
  }
  const rowData = [
    name,
    supervisor,
    date,
    leaveTypes[0] || "",
    leaveTypes[1] || "",
    startDate,
    endDate,
    null, //total days calculated in excel
    notes,
    "Pending", // Approval status column
  ];
  updateExcelRow(accessToken, siteId, fileId, emptyRow, rowData);

  //Send email notification
  sendEmailNotification(name, supervisor, date, accessToken, email);

  //If successful
  $message._show("success", "Submission Successful!");
  await new Promise((resolve) => setTimeout(resolve, 2000)); // Wait for 2 seconds
  window.location.href = "../html/mainpage.html"; // Redirect to main page
}

function validateInputs(
  name,
  supervisor,
  date,
  startDate,
  endDate,
  leaveTypes,
  otherReason
) {
  if (
    !name ||
    !supervisor ||
    !date ||
    !startDate ||
    !endDate ||
    leaveTypes.length === 0
  ) {
    alert("Please fill in all required fields.");
    return false;
  }

  if (leaveTypes.includes("other") && !otherReason.trim()) {
    alert("Please specify the 'Other' reason for leave.");
    return false;
  }

  return true;
}

async function sendEmailNotification(
  name,
  supervisor,
  date,
  accessToken,
  email
) {
  const adminEmail = "nicholas.hobden@rhinemechatronics.com";
  const subject = `Leave Application from ${name}`;
  const body = `
    <html>
      <body>
        <p>Good day ${supervisor},</p>
        <p>Kindly review my submitted leave application submitted on ${date}.</p>
        <p>Please log in to the leave application portal to view the details:</p>
        <p><a href="http://localhost:5500/index.html" target="_blank">Click here to view the leave application</a></p>
        <p>Thank you,</p>
        <p>${name}</p>
      </body>
    </html>
  `;

  // Email payload
  const emailPayload = {
    message: {
      subject: subject,
      body: {
        contentType: "HTML",
        content: body,
      },
      toRecipients: [
        {
          emailAddress: {
            address: adminEmail,
          },
        },
      ],
    },
  };

  // Send the email request to Microsoft Graph API
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users/${email}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(emailPayload),
      }
    );

    if (response.ok) {
      console.log("Email sent successfully!");
    } else {
      const error = await response.json();
      console.error("Error sending email:", error);
    }
  } catch (error) {
    console.error("Error in sending email via Graph API:", error);
  }
}
