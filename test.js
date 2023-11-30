var addon = require("bindings")("office");

const office = new addon.Office();
office.load("/usr/lib/libreoffice/program/");
for (let i = 0; i < 100; i++) {
  const start_time = Date.now();
  office.convert(
    "./test/Report.docx",
    "pdf",
    "./test/Report" + i + ".pdf"
  );
  console.log("took: %d",Date.now() - start_time)
}
office.close();
