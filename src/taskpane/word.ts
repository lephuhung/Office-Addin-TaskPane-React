/* global Word console */

export async function insertText(text: string) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function replaceSelection(text: string) {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function insertHeading(text: string, level: number = 1) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const para = body.insertParagraph(text, Word.InsertLocation.end);
      const styleName = `Heading ${Math.max(1, Math.min(3, Math.floor(level || 1)))}`;
      para.style = styleName;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
