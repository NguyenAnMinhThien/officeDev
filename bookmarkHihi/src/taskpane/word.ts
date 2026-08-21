/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runWord;
    document.getElementById("demo2").onclick = demo2;
  }
});

// Helper: Format date as yyyy-MM-dd
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

// // Helper: Extract review count from text using regex
// function extractReviewCount(text) {
//   const match = text.match(/\|(\d{1,3})\|/);
//   return match ? parseInt(match[1]) : 0;
// }

// // Helper: Get updated value with review count
// function getUpdatedReviewValue(text) {
//   const currentCount = extractReviewCount(text);
//   if (currentCount === 0) {
//     return "|1|";
//   } else {
//     return `|${currentCount + 1}|`;
//   }
// }

// // Helper: Remove old review marker from text
// function removeOldReviewMarker(text) {
//   return text.replace(/\s*\**\s*(\|\d+\|\s*)*/g, "").trim();
// }

// Helper: Extract review count from text using regex
function extractReviewCount(text) {
  const match = text.match(/\|(\d{1,3})\|/);
  return match ? parseInt(match[1]) : 0;
}

// Helper: Get updated value with review count
function getUpdatedReviewValue(text) {
  // Call the async function using the await keyword
  const currentCount = extractReviewCount(text);
  if (currentCount === 0) {
    return "|1|";
  } else {
    return `|${currentCount + 1}|`;
  }
}

// Helper: Remove old review marker from text
function removeOldReviewMarker(text) {
  return text.replace(/\s*\**\s*(\|\d+\|\s*)*/g, "").trim();
}

export async function runWord() {
  return Word.run(async (context) => {


    // 2. success get all paragraphs and extract text inside >
    // 1. Get the current active user selection
    const selection = context.document.getSelection();

    // 2. Fetch the paragraphs collection inside that selection
    const paragraphs = selection.paragraphs;

    // 3. CRITICAL: Load 'items' and the scalar properties you need for the loop
    // This resolves the "PropertyNotLoaded" error discussed earlier
    paragraphs.load("items/text");

    // 4. Synchronize with the Word document execution engine
    await context.sync();

    // 5. Safely loop through the retrieved paragraphs
    if (paragraphs.items.length > 0) {
      console.log(`Found ${paragraphs.items.length} paragraph(s) in selection:`);

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        console.log(`Paragraph ${i + 1}: ${para.text}`);
        // para.font.color = "#0000FF";
        para.font.color = "#00FF00";
        // para.font.color = "Automatic";
      }
    } else {
      console.log("No paragraphs found in the current selection.");
    }
    await context.sync();

    // 1. Get the current active user selection
// const selection = context.document.getSelection();

// // 2. Fetch the paragraphs collection inside that selection
// const paragraphs = selection.paragraphs;
// paragraphs.load("items/text, items/font/color");
// await context.sync();
// // 3. Set the font color directly on the collection using the White Hex string ("#FFFFFF")
// paragraphs.font.color = "#FFFFFF";

// // 4. Synchronize with the Word document engine to apply the change instantly
// await context.sync();
// console.log("All selected paragraphs have been set to white.");

    // success Get the first paragraph of the current selection >
    // const selection = context.document.getSelection();
    // const firstParagraph = selection.paragraphs.getFirst();
    // firstParagraph.load("text, font/color");
    // await context.sync();

    // // 1. succcess detec Red color >
    // const isRedFound = await detectRedTextInParagraph(firstParagraph, context);
    // console.log("Result: ", isRedFound);





  });
}

export async function demo2() {
  await Word.run(async (context) => {
    // // Get the first paragraph of the current selection
    // const selection = context.document.getSelection();
    // const firstParagraph = selection.paragraphs.getFirst();
    // firstParagraph.load("text, font/color");
    // await context.sync();

    // // 1. succcess detec Red color >
    // // const isRedFound = await detectRedTextInParagraph(firstParagraph, context);
    // // console.log("Result: ", isRedFound);

    // console.log("d");

    // Get current selection and expand to paragraph

        // 2. success get all paragraphs and extract text inside >
    // 1. Get the current active user selection
    const selection = context.document.getSelection();

    // 2. Fetch the paragraphs collection inside that selection
    const paragraphs = selection.paragraphs;

    // 3. CRITICAL: Load 'items' and the scalar properties you need for the loop
    // This resolves the "PropertyNotLoaded" error discussed earlier
    paragraphs.load("items/text, items/range, items/font/color");

    // 4. Synchronize with the Word document execution engine
    await context.sync();

    // const selection = context.document.getSelection();
    // selection.load("text");
    // await context.sync();

    // const paragraphs = context.document.body.paragraphs;
    // paragraphs.load("items/text,items/range");
    // await context.sync();

    // Find paragraph containing selection

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      const paraText = para.text;
      const hasRedText = await detectRedTextInParagraph(para, context);

      if (paraText.length > 4 && !hasRedText ) { 
        const newValue = getUpdatedReviewValue(paraText);
        const currentCount = extractReviewCount(paraText);

        if (currentCount === 0) {
          // No existing review marker, insert |1|
          para.insertText(`${newValue} `, Word.InsertLocation.start);
        } else {
          // Replace existing marker
          const oldMatch = paraText.match(/\s*\**\s*\|\d+\|\s*/);
          if (oldMatch) {
            const oldValue = oldMatch[0];


            // replace here
            // const searchResults = para
            //   .getRange()
            //   .getRange("Start")
            //   .getRange("End")
            //   .search(oldValue, { matchCase: false });

            // searchResults.load("items");
            // await context.sync();

            // if (searchResults.items.length > 0) {
            //   searchResults.items[0].insertText(`${newValue} `, Word.InsertLocation.replace);
            // }



        // 3. Search for the string target inside THIS paragraph
        const searchResults = para.search(oldValue, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();

        // 4. Check if at least one match exists, then replace only the first item (index 0)
        if (searchResults.items.length > 0) {
            const firstMatchRange = searchResults.items[0];
            firstMatchRange.insertText(newValue + " ", Word.InsertLocation.replace);
        }



          }
        }

        await context.sync();
        
      }
    }
  });
}


export async function greenParagraphs() {
  return Word.run(async (context) => {


    // 2. success get all paragraphs and extract text inside >
    // 1. Get the current active user selection
    const selection = context.document.getSelection();

    // 2. Fetch the paragraphs collection inside that selection
    const paragraphs = selection.paragraphs;

    // 3. CRITICAL: Load 'items' and the scalar properties you need for the loop
    // This resolves the "PropertyNotLoaded" error discussed earlier
    paragraphs.load("items/text");

    // 4. Synchronize with the Word document execution engine
    await context.sync();

    // 5. Safely loop through the retrieved paragraphs
    if (paragraphs.items.length > 0) {
      console.log(`Found ${paragraphs.items.length} paragraph(s) in selection:`);

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        console.log(`Paragraph ${i + 1}: ${para.text}`);
        // para.font.color = "#0000FF";
        para.font.color = "#00FF00";
        // para.font.color = "Automatic";
      }
    } else {
      console.log("No paragraphs found in the current selection.");
    }
    await context.sync();

    // 1. Get the current active user selection
// const selection = context.document.getSelection();

// // 2. Fetch the paragraphs collection inside that selection
// const paragraphs = selection.paragraphs;
// paragraphs.load("items/text, items/font/color");
// await context.sync();
// // 3. Set the font color directly on the collection using the White Hex string ("#FFFFFF")
// paragraphs.font.color = "#FFFFFF";

// // 4. Synchronize with the Word document engine to apply the change instantly
// await context.sync();
// console.log("All selected paragraphs have been set to white.");

    // success Get the first paragraph of the current selection >
    // const selection = context.document.getSelection();
    // const firstParagraph = selection.paragraphs.getFirst();
    // firstParagraph.load("text, font/color");
    // await context.sync();

    // // 1. succcess detec Red color >
    // const isRedFound = await detectRedTextInParagraph(firstParagraph, context);
    // console.log("Result: ", isRedFound);





  });
}


export async function markReviewed() {
  await Word.run(async (context) => {
    // // Get the first paragraph of the current selection
    // const selection = context.document.getSelection();
    // const firstParagraph = selection.paragraphs.getFirst();
    // firstParagraph.load("text, font/color");
    // await context.sync();

    // // 1. succcess detec Red color >
    // // const isRedFound = await detectRedTextInParagraph(firstParagraph, context);
    // // console.log("Result: ", isRedFound);

    // console.log("d");

    // Get current selection and expand to paragraph

        // 2. success get all paragraphs and extract text inside >
    // 1. Get the current active user selection
    const selection = context.document.getSelection();

    // 2. Fetch the paragraphs collection inside that selection
    const paragraphs = selection.paragraphs;

    // 3. CRITICAL: Load 'items' and the scalar properties you need for the loop
    // This resolves the "PropertyNotLoaded" error discussed earlier
    paragraphs.load("items/text, items/range, items/font/color");

    // 4. Synchronize with the Word document execution engine
    await context.sync();

    // const selection = context.document.getSelection();
    // selection.load("text");
    // await context.sync();

    // const paragraphs = context.document.body.paragraphs;
    // paragraphs.load("items/text,items/range");
    // await context.sync();

    // Find paragraph containing selection

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      const paraText = para.text;
      const hasRedText = await detectRedTextInParagraph(para, context);

      if (paraText.length > 4 && !hasRedText ) { 
        const newValue = getUpdatedReviewValue(paraText);
        const currentCount = extractReviewCount(paraText);

        if (currentCount === 0) {
          // No existing review marker, insert |1|
          para.insertText(`${newValue} `, Word.InsertLocation.start);
        } else {
          // Replace existing marker
          const oldMatch = paraText.match(/\s*\**\s*\|\d+\|\s*/);
          if (oldMatch) {
            const oldValue = oldMatch[0];


            // replace here
            // const searchResults = para
            //   .getRange()
            //   .getRange("Start")
            //   .getRange("End")
            //   .search(oldValue, { matchCase: false });

            // searchResults.load("items");
            // await context.sync();

            // if (searchResults.items.length > 0) {
            //   searchResults.items[0].insertText(`${newValue} `, Word.InsertLocation.replace);
            // }



        // 3. Search for the string target inside THIS paragraph
        const searchResults = para.search(oldValue, { matchCase: false, matchWholeWord: false });
        searchResults.load("items");
        await context.sync();

        // 4. Check if at least one match exists, then replace only the first item (index 0)
        if (searchResults.items.length > 0) {
            const firstMatchRange = searchResults.items[0];
            firstMatchRange.insertText(newValue + " ", Word.InsertLocation.replace);
        }



          }
        }

        await context.sync();
        
      }
    }
  });
}

/**
 * Detects if a paragraph contains text with an exact RGB red color.
 * @param {Word.Paragraph} paragraph - The target paragraph object.
 * @param {Word.RequestContext} context - The active request context.
 * @returns {Promise<boolean>} True if red text exists, otherwise false.
 */

// async function detectRedTextInParagraph(paragraph, context) {
//   // 1. Get all uniquely formatted text ranges (runs) inside the paragraph
//   // const textRanges = paragraph.textRanges; -> this function doesn't exist in current api
//   const textRanges = paragraph.getTextRanges();

//   // 2. Load the font color property and text for all segments
//   textRanges.load("items, items/font/color, text");
//   await context.sync();

//   let hasRedText = false;

//   // 3. Loop through the formatted segments instead of individual characters
//   for (let i = 0; i < textRanges.items.length; i++) {
//     const run = textRanges.items[i];

//     // Office JS returns hex colors (e.g., "#FF0000" for RGB 255, 0, 0)
//     if (run.font.color === "#FF0000") {
//       hasRedText = true;
//       break;
//     }
//   }

//   // 4. Output log mirroring your VBA Debug.Print behavior
//   if (hasRedText) {
//     paragraph.load("text");
//     await context.sync();
//     console.log(paragraph.text);
//   } else {
//     console.log("No red text found in this paragraph.");
//   }

//   return hasRedText;
// }

async function detectRedTextInParagraph(paragraph, context) {
  // 1. Pass an empty delimiter to split strictly by unique formatting chunks (runs)
  const textRanges = paragraph.getTextRanges([""]);

  // 2. CORRECTED: Remove the top-level "text" property from the load string.
  // It must look for items/text and items/font/color specifically.
  textRanges.load("items, items/font/color, items/text");
  await context.sync();

  let hasRedText = false;

  // 3. Loop through the formatted text segments safely
  for (let i = 0; i < textRanges.items.length; i++) {
    const run = textRanges.items[i];

    // Office JS returns hex colors (e.g., "#FF0000" for RGB 255, 0, 0)
    if (run.font.color === "#FF0000") {
      hasRedText = true;
      break;
    }
  }

  // 4. Output log mirroring your VBA Debug.Print behavior
  if (hasRedText) {
    paragraph.load("text");
    await context.sync();
    console.log(paragraph.text);
  } else {
    console.log("No red text found in this paragraph.");
  }

  return hasRedText;
}

// can select between from up to lower or lower to higher also
export async function SelectTextBetweenBookmarks() {
  return Word.run(async (context) => {
    // 1. Retrieve the bookmark range object using the document context
    const bookmarkRange = context.document.getBookmarkRangeOrNullObject("hehe");

    // 2. Load the range into the context execution queue
    bookmarkRange.load("address");

    await context.sync();

    // 3. Verify the bookmark exists in the document
    if (!bookmarkRange.isNullObject) {
      // 4. Get the starting point of the bookmark range
      // Valid string options for getRange() are "Whole", "Start", or "End"
      const startRange = bookmarkRange.getRange("Start");

      // Example action: Insert text or formatting exactly at the bookmark's start
      startRange.insertText("Text at Start: ", Word.InsertLocation.before);
      const selection = context.document.getSelection();
      const combinedRange = startRange.expandTo(selection);
      combinedRange.select();
      // startRange.load('start');
      await context.sync();
      // console.log(startRange.start);
    } else {
      console.log("The specified bookmark does not exist.");
    }

    await context.sync();
    // return;
  });
}

export async function putBookmark() {
  await Word.run(async (context) => {
    // 1. Get the current selection
    let selection = context.document.getSelection();

    // 2. Load the 'isEmpty' property
    selection.load("isEmpty");

    // 3. Sync with the Word document
    await context.sync();

    // // 4. Evaluate the boolean property
    // if (selection.isEmpty) {
    //   console.log("Selection is empty (only a flashing cursor).");
    // } else {
    //   console.log("Selection contains text or objects.");
    // }

    // Only add bookmark if there is actual selected content
    if (selection.isEmpty) {
      const now = new Date();
      // Create unique bookmark name: a_YYYYMMDD_HHMMSS format
      const year = 2100 - now.getFullYear();
      const month = String(100 - (now.getMonth() + 1)).padStart(2, "0");
      const day = String(100 - now.getDate()).padStart(2, "0");
      const hours = String(100 - now.getHours()).padStart(2, "0");
      const minutes = String(100 - now.getMinutes()).padStart(2, "0");
      const seconds = String(100 - now.getSeconds()).padStart(2, "0");

      const bookmarkName = `a_${year}${month}${day}_${hours}${minutes}${seconds}`;
      console.log("Selection empty", bookmarkName);

      selection.insertBookmark(bookmarkName);
    } else {
      // If no selection, create bookmark at current position
      const now = new Date();
      const year = 2100 - now.getFullYear();
      const month = String(100 - (now.getMonth() + 1)).padStart(2, "0");
      const day = String(100 - now.getDate()).padStart(2, "0");
      const hours = String(100 - now.getHours()).padStart(2, "0");
      const minutes = String(100 - now.getMinutes()).padStart(2, "0");
      const seconds = String(100 - now.getSeconds()).padStart(2, "0");

      const bookmarkName = `a_${year}${month}${day}_${hours}${minutes}${seconds}`;
      console.log("Selection NOT empty", bookmarkName);
      selection.insertBookmark(bookmarkName);
    }
    await context.sync();
  });
}

export async function jumpToBookmark() {
  await Word.run(async (context) => {
    let doc = context.document;
    let bookmarkRange = doc.getBookmarkRangeOrNullObject("hehe");

    // Load properties to check existence and read data
    bookmarkRange.load(["text", "isNullObject"]);
    await context.sync();

    if (bookmarkRange.isNullObject) {
      console.log("Bookmark not found.");
    } else {
      console.log("Bookmark text contents: " + bookmarkRange.text);
      bookmarkRange.getRange("Start").select();

      // Example: Update text within the bookmark
      // bookmarkRange.insertText("Updated content here.", Word.InsertLocation.replace);
      await context.sync();
    }

    // const bookmarks = context.document.bookmarks;
    //       bookmarks.load("items");
    //       return context.sync().then(() => {
    //         const heheBookmark = bookmarks.getItem("hehe");
    //         if (heheBookmark) {
    //           heheBookmark.getRange("Start").select();
    //           return context.sync();
    //         } else {
    //           alert("Bookmark 'hehe' not found");
    //         }
    //       });

    // const bookmarks = context.document.bookmarks;
    // bookmarks.load("items");
    // await context.sync();

    // 1. Reference the bookmarks collection
    // const bookmarks = context.document.bookmarks;

    // 2. Queue the command to load the collection properties (specifically 'items')
    // You can also load specific properties per bookmark like 'name'
    // bookmarks.load("items");

    // // 3. Synchronize the state with the Word document
    // await context.sync();

    // // 4. Iterate through the loaded bookmarks
    // if (bookmarks.items.length === 0) {
    //   console.log("No bookmarks found in the document.");
    // } else {
    //   bookmarks.items.forEach((bookmark) => {
    //     console.log(`Bookmark Name: ${bookmark.name}`);
    //   });
    // }

    // document.getElementById("huhu").textContent = bookmarks[0];
    // try {
    //   // Try to get the bookmark
    //   const heheBookmark = bookmarks.getItem("hehe");

    //   // Load its range
    //   heheBookmark.load("name");
    //   await context.sync();

    //   // Move caret to the bookmark's start
    //   const caret = heheBookmark.getRange("Start");
    //   caret.select(); // zero-length → real cursor appears

    //   await context.sync();
    // } catch (e) {
    //   // Bookmark not found
    //   console.log("Bookmark 'hehe' not found");
    //   alert("Bookmark 'hehe' not found");
    // }
  });
}

export async function goHome() {
  await Word.run(async (context) => {
    await addBookmarkHere();

    // const newPara = startRange.insertParagraph("", Word.InsertLocation.start);
    const newPara = context.document.body.insertParagraph("", Word.InsertLocation.start);
    // context.document.body.getRange("End").select();
    const caret = newPara.getRange("Start");

    // Select the zero-length range → but cursor appears there
    caret.select();

    return context.sync();
  });
}

export async function addBookmarkHere() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();

    // Load the range so we can use it
    selection.load("text");
    await context.sync();

    // Insert bookmark at the selection
    selection.insertBookmark("hehe");

    await context.sync();
  });
}

export async function goEnd() {
  await Word.run(async (context) => {
    await addBookmarkHere();
    const endRange = context.document.body.getRange("End");
    // endRange.insertBookmark("hehe");
    const newPara = endRange.insertParagraph("", Word.InsertLocation.after);
    // context.document.body.getRange("End").select();
    const caret = newPara.getRange("End");

    // Select the zero-length range → cursor appears there
    caret.select();
    // // 2. Check compatibility and pull focus from the task pane to the document
    // if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.4")) {
    //   document.getElementById("huhu").textContent = "Your updated text here.";
    //   context.document.activeWindow.setFocus();
    // }
    // document.getElementById("huhu").textContent = "Your updated";
    // context.document.activeWindow.setFocus();   //not work because this is not supported
    // newPara.select();
    return context.sync();
  });
}

// export async function insertDatetime() {

//   return Word.run((context) => {
//     const now = new Date();
//     const dateStr = formatDate(now);
//     const selection = context.document.getSelection();
//     selection.insertText(dateStr, Word.InsertLocation.replace);
//     return context.sync();
//   });
// }
