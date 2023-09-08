const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
  path.resolve(__dirname, "input.docx"),
  "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
  fullName: "Скворцова А.Б.",
  classes: [
    {
      number: 1,
      nameClass: "6Б",
      generalCaunt: 34,
      writed: 29,
      five: 7,
      four: 14,
      three: 7,
      two: 1,
      procentQuality: 72.41,
      procentProgress: 23.13,
      GPA: 3.93,
    },
    {
      number: 2,
      nameClass: "7Д",
      generalCaunt: 25,
      writed: 20,
      five: 2,
      four: 4,
      three: 11,
      two: 3,
      procentQuality: 30,
      procentProgress: 55,
      GPA: 3.25,
    },
    {
      number: 3,
      nameClass: "7Г",
      generalCaunt: 29,
      writed: 24,
      five: 3,
      four: 5,
      three: 10,
      two: 6,
      procentQuality: 33.3,
      procentProgress: 41.66,
      GPA: 3.2,
    },
  ],
});

const buf = doc.getZip().generate({
  type: "nodebuffer",
  // compression: DEFLATE adds a compression step.
  // For a 50MB output document, expect 500ms additional CPU time
  compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a
// file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
