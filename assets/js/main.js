import { getExcelData, SlideShowBG } from "./exports.js";
import { canUse } from "./exports.js";

(function () {
  "use strict"; //strict js to jelp reduce accidental errors like undeclared variables

  var $body = document.querySelector("body"); //selects body tag from html

  // Methods/polyfills. These make sure that the methods used are supported by older browsers

  // classList | (c) @remy | github.com/remy/polyfills | rem.mit-license.org
  !(function () {
    function t(t) {
      this.el = t;
      for (
        var n = t.className.replace(/^\s+|\s+$/g, "").split(/\s+/), i = 0;
        i < n.length;
        i++
      )
        e.call(this, n[i]);
    }
    function n(t, n, i) {
      Object.defineProperty
        ? Object.defineProperty(t, n, { get: i })
        : t.__defineGetter__(n, i);
    }
    if (
      !(
        "undefined" == typeof window.Element ||
        "classList" in document.documentElement
      )
    ) {
      var i = Array.prototype,
        e = i.push,
        s = i.splice,
        o = i.join;
      (t.prototype = {
        add: function (t) {
          this.contains(t) ||
            (e.call(this, t), (this.el.className = this.toString()));
        },
        contains: function (t) {
          return -1 != this.el.className.indexOf(t);
        },
        item: function (t) {
          return this[t] || null;
        },
        remove: function (t) {
          if (this.contains(t)) {
            for (var n = 0; n < this.length && this[n] != t; n++);
            s.call(this, n, 1), (this.el.className = this.toString());
          }
        },
        toString: function () {
          return o.call(this, " ");
        },
        toggle: function (t) {
          return (
            this.contains(t) ? this.remove(t) : this.add(t), this.contains(t)
          );
        },
      }),
        (window.DOMTokenList = t),
        n(Element.prototype, "classList", function () {
          return new t(this);
        });
    }
  })();

  // canUse

  // window.addEventListener. Ensure that window.addEventListener is available in older browsers
  (function () {
    if ("addEventListener" in window) return;
    window.addEventListener = function (type, f) {
      window.attachEvent("on" + type, f);
    };
  })();

  // Play initial animations on page load.
  window.addEventListener("load", function () {
    window.setTimeout(function () {
      $body.classList.remove("is-preload");
    }, 100);
  });

  //Start slideshow background
  SlideShowBG(0);

  // Auto-login if "Remember Me" was previously selected
  (function () {
    const msalConfig = {
      auth: {
        clientId: "cf1a04b4-d42a-4e18-9ec3-18ec43bfeaf9",
        authority:
          "https://login.microsoftonline.com/14e5adcf-4446-4cee-bac9-e293492fa769",
        redirectUri: "http://localhost:5500/index.html",
      },
      cache: {
        cacheLocation: "localStorage", // Try localStorage where "Remember Me" would have saved it
        storeAuthStateInCookie: true,
      },
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);

      msalInstance
        .acquireTokenSilent({
          account: accounts[0],
          scopes: [
            "User.Read",
            "Mail.Read",
            "Sites.Read.All",
            "Files.Read.All",
          ],
        })
        .then((tokenResponse) => {
          const accessToken = tokenResponse.accessToken;

          return fetch_access(null, msalInstance, accessToken);
        })
        .catch((error) => {
          console.warn("Silent auto-login failed:", error);
          // Do nothing, fallback to manual login
        });
    }
  })();

  // login Form.
  (function () {
    const $form = document.querySelector("#login-form"),
      $submit = document.querySelector('#login-form input[type="submit"]');

    if (!("addEventListener" in $form)) return;

    const $message = document.createElement("span");
    $message.classList.add("message");
    $form.appendChild($message);

    $message._show = function (type, text) {
      $message.innerHTML = text;
      $message.classList.add(type, "visible");
      setTimeout(() => $message._hide(), 3000);
    };

    $message._hide = function () {
      $message.classList.remove("visible");
    };

    function handleLoginSubmit(event) {
      event.preventDefault();
      $message._hide();
      $submit.disabled = true;

      const rememberMe = document.getElementById("remember-me").checked;
      const storageType = rememberMe ? "localStorage" : "sessionStorage";

      const msalConfig = {
        auth: {
          clientId: "cf1a04b4-d42a-4e18-9ec3-18ec43bfeaf9",
          authority:
            "https://login.microsoftonline.com/14e5adcf-4446-4cee-bac9-e293492fa769",
          redirectUri: "http://localhost:5500/index.html",
        },
        cache: {
          cacheLocation: storageType,
          storeAuthStateInCookie: true,
        },
      };

      const msalInstance = new msal.PublicClientApplication(msalConfig);
      const accounts = msalInstance.getAllAccounts();

      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);

        msalInstance
          .acquireTokenSilent({
            account: accounts[0],
            scopes: [
              "User.Read",
              "Mail.Read",
              "Sites.Read.All",
              "Files.Read.All",
            ],
          })
          .then((tokenResponse) => {
            const accessToken = tokenResponse.accessToken;

            return fetch_access($message, msalInstance, accessToken);
          })
          .catch((error) => {
            console.error("Silent token acquisition failed", error);
            $message._show("failure", "Session expired. Please login again.");
            $submit.disabled = false;
          });
      } else {
        // Login popup if no account cached
        const loginRequest = {
          scopes: [
            "User.Read",
            "Mail.Read",
            "Sites.Read.All",
            "Files.Read.All",
            "Sites.ReadWrite.All",
          ],
        };

        msalInstance
          .loginPopup(loginRequest)
          .then((response) => {
            const accessToken = response.accessToken;
            msalInstance.setActiveAccount(response.account);

            return fetch_access($message, msalInstance, accessToken);
          })
          .catch((error) => {
            console.error("Login failed", error);
            $message._show("failure", "Login failed");
            $submit.disabled = false;
          });
      }
    }

    $form.addEventListener("submit", handleLoginSubmit);
  })();
})();

// Function to fetch user profile and access token
function fetch_access($message, msalInstance, accessToken) {
  fetch("https://graph.microsoft.com/v1.0/me", {
    method: "GET",
    headers: { Authorization: "Bearer " + accessToken },
  })
    .then((response) => response.json())
    .then((data) => {
      sessionStorage.setItem("userProfile", JSON.stringify(data));
      sessionStorage.setItem("accessToken", accessToken);
      sessionStorage.setItem("email", data.mail);

      return msalInstance.acquireTokenSilent({
        scopes: ["Sites.Read.All", "Files.Read.All"],
      });
    })
    .then(async (tokenResponse) => {
      sessionStorage.setItem("sharepointToken", tokenResponse.accessToken);
      if ($message != null) {
        $message._show("success", "Login Successful!");
      }

      const siteId =
        "netorg7968809.sharepoint.com,d6ef5094-875f-47d7-93c4-43ae171a04ff,883a8121-0374-49f4-9476-2d3b9a1cb38a";

      const fileId = "012LJMUY6BHXDWVGWPI5DIT3YPOFVODUTI";
      const excelData = await getExcelData(
        accessToken,
        siteId,
        fileId,
        "Admins"
      );
      for (let i = 1; i < excelData.length; i++) {
        if (tokenResponse.account.name === excelData[i][0]) {
          window.location.href =
            "http://localhost:5500/html/leave_application.html";
          break;
        } else {
        }
      }
      window.location.href = "http://localhost:5500/html/mainpage.html";
    });
}
