/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
  }

// Assign event handlers and other initialization logic.
document.getElementById("insert-paragraph").onclick = insertParagraph;
document.getElementById("apply-style").onclick = applyStyle;
document.getElementById("apply-custom-style").onclick = applyCustomStyle;
document.getElementById("change-font").onclick = changeFont;
document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
  }
});

function insertParagraph() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.body;
      docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                        "Start");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function applyStyle() {
  Word.run(function (context) {

      // TODO1: Queue commands to style text.
      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.Style.intenseReference;
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function applyCustomStyle() {
  Word.run(function (context) {

      // TODO1: Queue commands to apply the custom style.
      var lastParagraph = context.document.body.paragraphs.getLast();
      lastParagraph.style = "MyCustomStyle";
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function changeFont() {
  Word.run(function (context) {

      // TODO1: Queue commands to apply a different font.
      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18
    });
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTextIntoRange() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert text into a selected range.
      var doc = context.document;
      var originalRange = doc.getSelection();
      originalRange.insertText(" (C2R)", "End");
      // TODO2: Load the text of the range and sync so that the
      //        current range text can be read.
      
      // TODO3: Queue commands to repeat the text of the original
      //        range at the end of the document.
      doc.body.insertParagraph("Original range: " + originalRange.text, "End");
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}