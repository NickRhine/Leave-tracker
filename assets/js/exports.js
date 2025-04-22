"use strict";
var $body = document.querySelector("body");
// Slideshow Background.
export function SlideShowBG(sub_dir) {
  // Settings. Every 6 seconds the background image changes
  var settings = {
    // Images (in the format of 'url': 'alignment').
    images: {},
    // Delay.
    delay: 6000,
  };
  if (sub_dir === 1) {
    settings.images = {
      "../images/bg01.jpg": "center",
      "../images/bg02.jpg": "center",
      "../images/bg03.jpg": "center",
    };
  } else {
    settings.images = {
      "images/bg01.jpg": "center",
      "images/bg02.jpg": "center",
      "images/bg03.jpg": "center",
    };
  }
  // Creates div container for the background and inserts into body
  // Vars.
  var pos = 0,
    lastPos = 0,
    $wrapper,
    $bgs = [],
    $bg,
    k,
    v;

  // Create BG wrapper, BGs.
  $wrapper = document.createElement("div");
  $wrapper.id = "bg";
  $body.appendChild($wrapper);

  // Loops through the images and creates divs for each image and stores them in an array
  for (k in settings.images) {
    // Create BG.
    $bg = document.createElement("div");
    $bg.style.backgroundImage = 'url("' + k + '")';
    $bg.style.backgroundPosition = settings.images[k];
    $wrapper.appendChild($bg);

    // Add it to array.
    $bgs.push($bg);
  }

  // Main loop. This function is called every 6 seconds to change the background image
  // Main loop.
  $bgs[pos].classList.add("visible");
  $bgs[pos].classList.add("top");

  // Bail if we only have a single BG or the client doesn't support transitions.
  if ($bgs.length == 1 || !canUse("transition")) return;

  window.setInterval(function () {
    lastPos = pos;
    pos++;

    // Wrap to beginning if necessary.
    if (pos >= $bgs.length) pos = 0;

    // Swap top images.
    $bgs[lastPos].classList.remove("top");
    $bgs[pos].classList.add("visible");
    $bgs[pos].classList.add("top");

    // Hide last image after a short delay.
    window.setTimeout(function () {
      $bgs[lastPos].classList.remove("visible");
    }, settings.delay / 2);
  }, settings.delay);
}

export function canUse(p) {
  if (!window._canUse) window._canUse = document.createElement("div");
  var e = window._canUse.style,
    up = p.charAt(0).toUpperCase() + p.slice(1);
  return (
    p in e ||
    "Moz" + up in e ||
    "Webkit" + up in e ||
    "O" + up in e ||
    "ms" + up in e
  );
}

export async function getExcelData(accessToken, siteId, fileId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('Data')/usedRange`;

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

// Excel stores dates in serial format where 1 corresponds to 1899-11-30
export function excelSerialDateToJSDate(serial) {
  const excelEpoch = new Date(1899, 11, 30);
  return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}
