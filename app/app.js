// // Common app functionality

// var app = (function () {
//     'use strict';

//     var app = {};

//     // Common initialization function (to be called from each page)
//     app.initialize = function () {
//         $('body').append(
// 			'<div id="notification-message">' +
// 				'<div class="padding">' +
// 					'<div id="notification-message-close"></div>' +
// 					'<div id="notification-message-header"></div>' +
// 					'<div id="notification-message-body"></div>' +
// 				'</div>' +
// 			'</div>');

//         $('#notification-message-close').click(function () {
//             $('#notification-message').hide();
//         });


//         // After initialization, expose a common notification function
//         app.showNotification = function (header, text) {
//             $('#notification-message-header').text(header);
//             $('#notification-message-body').text(text);
//             $('#notification-message').slideDown('fast');
//         };
//     };

//     return app;
// })();


// (function () {
//     'use strict';

//     // The initialize function must be run each time a new page is loaded
//     Office.initialize = function (reason) {
//         $(document).ready(function () {
//             app.initialize();

//             //$('#get-data-from-selection').click(getDataFromSelection);
//         });
//     };

//     // Reads data from current document selection and displays a notification
//     function getDataFromSelection() {
//         Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
//             function (result) {
//                 if (result.status === Office.AsyncResultStatus.Succeeded) {
//                     app.showNotification('The selected text is:', '"' + result.value + '"');
//                 } else {
//                     app.showNotification('Error:', result.error.message);
//                 }
//             }
//         );
//     }
// })();