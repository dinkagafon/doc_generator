const fs = require("fs");
const docx = require("docx");
const { HeadingLevel, Table, TableRow, TableCell } = require("docx");

const { Document, Packer, Paragraph, TextRun } = docx;

const RESULT_DOC_FILE_NAME = "example";
const FOLDER_NAME = "folder";
const FOLDER_PATH = "../";

const exclude = [
  "node_modules",
  ".git",
  "README.md",
  "yarn.lock",
  "database.db",
  "tests",
  "img",
  "fonts",
];

const isFile = (fileName) => {
  return fs.lstatSync(fileName).isFile();
};

const titles = [];

const generateTable = (path, fileName) => {
  const array = fs
    .readFileSync(path + fileName)
    .toString()
    .split("\n");
  const codeParagraphs = [];
  for (let i of array) {
    codeParagraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: i,
            font: "Courier New",
            size: 24,
            bold: true,
          }),
        ],
      })
    );
  }
  const table = new Table({
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: fileName,
                    font: "Courier New",
                    size: 24,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            children: codeParagraphs,
          }),
        ],
      }),
    ],
  });

  return table;
};

const openFolder = (path, folderName, deep) => {
  const folderPath = path + folderName;
  titles.push(
    new Paragraph({
      text: folderName,
      heading: HeadingLevel["HEADING_" + deep],
    })
  );

  fs.readdirSync(folderPath).forEach((name) => {
    if (exclude.includes(name) || name.endsWith(".png")) {
      return;
    }
    if (isFile(folderPath + "/" + name)) {
      titles.push(
        new Paragraph({
          text: name,
          heading: HeadingLevel["HEADING_" + (deep + 1)],
        })
      );
      const table = generateTable(folderPath + "/", name);
      titles.push(table);
    } else {
      openFolder(folderPath + "/", name, deep + 1);
    }
  });
};

openFolder(FOLDER_PATH, FOLDER_NAME, 1);

const doc = new Document({
  sections: [
    {
      properties: {},
      children: titles,
    },
  ],
});

// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(`${RESULT_DOC_FILE_NAME}.docx`, buffer);
});
