async function fetchPendingRequests() {
  const accessToken = await getAccessToken();
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/sites/{site-id}/drives/{drive-id}/items/{excel-id}/workbook/worksheets/{worksheet-name}/tables/{table-name}/rows",
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );
  const data = await response.json();

  const requests = data.value;
  const container = document.getElementById("leave-requests");
  container.innerHTML = "";

  requests.forEach((row, index) => {
    const fields = row.values[0]; // adjust if your Excel format is different
    const [name, supervisor, startDate, endDate, leaveType, notes, status] =
      fields;

    if (status === "Pending") {
      const div = document.createElement("div");
      div.innerHTML = `
          <p><strong>${name}</strong> (${leaveType})</p>
          <p>${startDate} to ${endDate}</p>
          <p>Notes: ${notes}</p>
          <button onclick="updateStatus(${index}, 'Approved')">Approve</button>
          <button onclick="updateStatus(${index}, 'Denied')">Deny</button>
          <hr/>
        `;
      container.appendChild(div);
    }
  });
}

async function updateStatus(rowIndex, newStatus) {
  const accessToken = await getAccessToken();
  await fetch(
    `https://graph.microsoft.com/v1.0/sites/{site-id}/drives/{drive-id}/items/{excel-id}/workbook/worksheets/{worksheet-name}/tables/{table-name}/rows/${rowIndex}`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        values: [[null, null, null, null, null, null, newStatus]],
      }),
    }
  );

  alert("Status updated.");
  fetchPendingRequests();
}

document.addEventListener("DOMContentLoaded", fetchPendingRequests);
