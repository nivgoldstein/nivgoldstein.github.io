var dialog;

function allowToSend(asyncResult) {
  return () => {
    asyncResult.asyncContext.completed({ allowEvent: true });
  }
}

function notAllowedToSend(asyncResult) {
  return () => {
    asyncResult.asyncContext.completed({ allowEvent: false });
  }
}

function showDialog(approveFn, cancelFn) {
  var dialogUrl = 'https://' + location.host + '/dialog.html'
  Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20 },
    function (asyncResult) {
      dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => processMessage(arg, approveFn, cancelFn));
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => eventHandler(arg, cancelFn));

    }
  );
}

function eventHandler(arg, cancelFn) {
  // In addition to general system errors, there are 2 specific errors 
  // and one event that you can handle individually.
  // switch (arg.error) {
  //   case 12002:
  //     showNotification("Cannot load URL, no such page or bad URL syntax.");
  //     break;
  //   case 12003:
  //     showNotification("HTTPS is required.");
  //     break;
  //   case 12006:
  //     // The dialog was closed, typically because the user the pressed X button.
  //     showNotification("Dialog closed by user");
  //     break;
  //   default:
  //     showNotification("Undefined error in dialog window");
  //     break;
  // }
  if (arg.error) {
    cancelFn()
  }
}

function processMessage(arg, approveFn, cancelFn) {
  var messageFromDialog = JSON.parse(arg.message);
  console.log(messageFromDialog)
  if (messageFromDialog.result === 'yes') {
    // do something
    approveFn()
  } else if (messageFromDialog.result === 'no') {
    // do something
    cancelFn()
  } else {
    // closed by something else
    cancelFn()
  }
  dialog.close();
}

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

// Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
// <param name="event">MessageSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
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
          // get the list of people
          console.log(r, "In fetch delay")
          showDialog(allowToSend(asyncResult), notAllowedToSend(asyncResult))
          // asyncResult.asyncContext.completed({ allowEvent: false });
        }
      )

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
