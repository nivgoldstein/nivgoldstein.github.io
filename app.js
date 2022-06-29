/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var mailboxItem;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
  mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}

// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
  shouldChangeSubjectOnSend(event);
}

// function openDialog() {
//   const dialog = document.createElement('dialog')
//   const card = document.getElementById('ReadingPaneContainerId')
//   const approveBtn = document.createElement('button')
//   const cancleBtn = document.createElement('button')
//   approveBtn.innerText = 'Approve'
//   cancleBtn.innerText = 'Cancel'

//   approveBtn.addEventListener('click', () => {
//     // callback here for approve
//     alert('click')
//   })

//   dialog.appendChild(approveBtn)
//   dialog.appendChild(cancleBtn)
//   dialog.style.width = '250px'
//   dialog.style.height = '80px'
//   dialog.style.backgroundColor = 'red'
//   dialog.open = true
//   dialog.style.top = '50%'
//   dialog.style.transform = 'translateY(-50%)'
//   dialog.style.display = 'flex'
//   dialog.style.justifyContent = 'space-around'
//   dialog.style.alignItems = 'center'
//   document.body.appendChild(dialog)
//   console.log('open dialog 3', document.body)

// }

// Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
// <param name="event">MessageSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
  console.log(event)
  MessageBox.Show("Do you like icecream?", "Questionaire", MessageBoxButtons.YesNo,
    MessageBoxIcons.Question, false, null, function (buttonFirst) {
      /** @type {string} */
      var iceCream = (buttonFirst == "Yes" ? "do" : "dont");
      MessageBox.UpdateMessage("Do you like Jelly Beans?", function (buttonSecond) {
        /** @type {string} */
        var jellyBeans = (buttonSecond == "Yes" ? "do" : "dont");
        MessageBox.UpdateMessage("Do you like Kit Kat bars?", function (buttonThird) {
          /** type {string} */
          var kitkat = (buttonThird == "Yes" ? "do" : "dont");
          MessageBox.CloseDialogAsync(function () {
            Alert.Show("You said you " + iceCream + " like ice cream, you " +
              jellyBeans + " like jelly beans, and you " +
              kitkat + " like kit kat bars.");
          });
        });
      });
    }, true);
  mailboxItem.subject.getAsync(
    { asyncContext: event },
    function (asyncResult) {
      // addCCOnSend(asyncResult.asyncContext);
      //console.log(asyncResult.value);
      // Match string.
      console.log("V5")

      fetch("https://httpbin.org/delay/5").then(
        r => {
          return r.json()
        }
      ).then(
        r => {
          console.log("In fetch delay")
          asyncResult.asyncContext.completed({ allowEvent: false });
        }
      )
      console.log("After fetch delay")

      // setTimeout(() => {
      //     console.log("In Timeout")
      //     asyncResult.asyncContext.completed({ allowEvent: true });
      //
      // }, 5000)
      // console.log("After Timeout")

      // asyncResult.asyncContext.completed({ allowEvent: false });

      // mailboxItem.notificationMessages.addAsync(
      //     key='A',
      //     JSONmessage={
      //         type: 'errorMessage',
      //         message: 'Test message 4',
      //         action: [
      //             {
      //                 actionText: "A1",
      //                 actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
      //                 commandId: "msgComposeOpenPaneButton",
      //                 contextData: JSON.stringify({a: "aValue", b: "bValue"}),
      //             },
      //         ],
      //     },
      //     callback=(result) => {
      //         console.log("In")
      //     }
      // );
      //
      // console.log("Out")
      // asyncResult.asyncContext.completed({ allowEvent: false });






      // var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
      // // Add [Checked]: to subject line.
      // subject = '[Checked]: ' + asyncResult.value;
      //
      // // Check if a string is blank, null or undefined.
      // // If yes, block send and display information bar to notify sender to add a subject.
      // if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
      //     mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
      //     asyncResult.asyncContext.completed({ allowEvent: false });
      // }
      // else {
      //     // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
      //     if (!checkSubject) {
      //         subjectOnSendChange(subject, asyncResult.asyncContext);
      //         //console.log(checkSubject);
      //     }
      //     else {
      //         // Allow send.
      //         asyncResult.asyncContext.completed({ allowEvent: true });
      //     }
      // }

    }
  )
}

// Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">MessageSend event passed from calling function</param>
function addCCOnSend(event) {
  mailboxItem.cc.setAsync(['ngoldstein@ironscales.com', 'nivgoldstein123@gmail.com'], { asyncContext: event });
}

// Check if the subject should be changed. If it is already changed allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">MessageSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
  mailboxItem.subject.setAsync(
    subject,
    { asyncContext: event },
    function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
      }
      else {
        // Allow send.
        asyncResult.asyncContext.completed({ allowEvent: true });
      }

    });
}

// Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
// <param name="asyncResult">MessageSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
  var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
  var wordExpression = listOfBlockedWords.join('|');

  // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
  // i to perform case-insensitive search.
  var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
  var checkBody = regexCheck.test(asyncResult.value);

  if (checkBody) {
    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
    // Block send.
    asyncResult.asyncContext.completed({ allowEvent: false });
  }
  else {

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
  }
}
