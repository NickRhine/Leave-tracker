import { SlideShowBG } from "./exports.js";
import { canUse } from "./exports.js";

(function () {
  "use strict"; //strict js to jelp reduce accidental errors like undeclared variables

  var $body = document.querySelector("body"); //selects body tag from html

  const msalConfig = {
    auth: {
      clientId: "cf1a04b4-d42a-4e18-9ec3-18ec43bfeaf9",
      authority:
        "https://login.microsoftonline.com/14e5adcf-4446-4cee-bac9-e293492fa769",
      redirectUri: "http://127.0.0.1:5500/html/mainpage.html",
    },
    cache: {
      cacheLocation: "localStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: true, // Set this to "true" if you want to store the auth state in a cookie
    },
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);

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

  SlideShowBG(0);

  // login Form.
  (function () {
    // Vars.
    var $form = document.querySelectorAll("#login-form")[0],
      $submit = document.querySelectorAll(
        '#login-form input[type="submit"]'
      )[0],
      $message;

    // Bail if addEventListener isn't supported.
    if (!("addEventListener" in $form)) return;

    // Message.
    $message = document.createElement("span");
    $message.classList.add("message");
    $form.appendChild($message);

    $message._show = function (type, text) {
      $message.innerHTML = text;
      $message.classList.add(type);
      $message.classList.add("visible");

      setTimeout(function () {
        $message._hide();
      }, 3000);
    };

    $message._hide = function () {
      $message.classList.remove("visible");
    };

    // Events.
    $form.addEventListener("submit", function (event) {
      // event.stopPropagation();
      event.preventDefault();

      // Hide message.
      $message._hide();

      var email = document.querySelector("#email").value.trim();
      var password = document.querySelector("#password").value.trim();

      document.querySelector("#email").classList.remove("input-error");
      document.querySelector("#password").classList.remove("input-error");

      if (password === "") {
        document.querySelector("#password").classList.add("input-error");
        $message._show("failure", "Password cannot be empty");
        $submit.disabled = false;
        return;
      }

      if (email === "" || !email.includes("@")) {
        document.querySelector("#email").classList.add("input-error");
        $message._show("failure", "Please enter a valid email");
        $submit.disabled = false;
        return;
      }
      // Disable submit.
      $submit.disabled = true;

      // Microsoft login
      const loginRequest = {
        scopes: ["User.Read", "Mail.Read"], // Define the necessary Microsoft Graph API permissions
      };

      msalInstance
        .loginPopup(loginRequest)
        .then((response) => {
          console.log("Login Successful", response);

          const accessToken = response.accessToken;
          getUserProfile(accessToken);

          $message._show("success", "Login Successful!");
        })
        .catch((error) => {
          console.error("Login failed", error);
          $message._show("failure", "Login failed");
          $submit.disabled = false;
        });
      // Process form.
      // Note: Doesn't actually do anything yet (other than report back with a "thank you"),
      // but there's enough here to piece together a working AJAX submission call that does.
      // setTimeout(function () {
      //   // Reset form.
      //   $form.reset();

      //   // Enable submit.
      //   $submit.disabled = false;
      //   $message.classList.remove("failure");

      //   // Show message.
      //   $message._show = function (type, text) {
      //     $message.innerHTML = text;
      //     $message.classList.add(type);
      //     $message.classList.add("visible");

      //     setTimeout(function () {
      //       $message._hide();
      //     }, 3000);
      //   };
      //   $message._show("success", "Login Successful!");
      // }, 750);
    });

    function getUserProfile(token) {
      fetch("https://graph.microsoft.com/v1.0/me", {
        method: "GET",
        headers: {
          Authorization: "Bearer " + token,
        },
      })
        .then((response) => response.json())
        .then((data) => {
          console.log("User Profile:", data);
          // You can now use the user profile data as needed
          // For example, you can display the user's name or email on the page
          document.querySelector(
            "#user-name"
          ).textContent = `Welcome, ${data.displayName}`;
        })
        .catch((error) => {
          console.error("Error fetching user profile:", error);
        });
    }
  })();
})();

// document.addEventListener("DOMContentLoaded", function () {
//   let glow = document.createElement("div");
//   glow.classList.add("cursor-glow");
//   document.body.appendChild(glow);

//   document.addEventListener("mousemove", function (event) {
//     let x = event.clientX - 50;
//     let y = event.clientY - 50;
//     glow.style.transform = `translate(${x}px, ${y}px)`;
//   });
// });
