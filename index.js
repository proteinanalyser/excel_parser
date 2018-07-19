var express = require('express');
var bodyParser = require("body-parser");
var path = require("path");
var app = express();
var multer = require('multer')
var XLSX = require('xlsx')
var fs = require("fs")
const storage = multer.diskStorage({
  destination: "./uploads",
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  }
});
const upload = multer({
  storage: storage,
  fileFilter: function (req, file, cb) {
    checkFileType(file, cb);
  }
}).single("myfile");

function checkFileType(file, cb) {
  let filetypes = ".xlsx";
  let extname = path.extname(file.originalname).toLowerCase() == filetypes;
  let mimetype = file.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  if (path.extname(file.originalname) != ".xlsx") {
    cb("The file you uploaded has a wrong file extension! The file must be .xlsx");
  } else if (file.originalname != "template.xlsx") {
    cb("Only upload the template file provided by this platform!");
  } else if (mimetype && extname) {
    return cb(null, true);
  }
}
// Add headers
app.use(function (req, res, next) {

  // Website you wish to allow to connect
  res.setHeader('Access-Control-Allow-Origin', 'http://agirar.pt');
  // res.setHeader('Access-Control-Allow-Origin', 'http://localhost:8000');

  // Request methods you wish to allow
  res.setHeader('Access-Control-Allow-Methods', 'POST');

  // Request headers you wish to allow
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');
  res.setHeader('Access-Control-Allow-Credentials', true);

  // Pass to next layer of middleware
  next();
});

app.set("port", (process.env.PORT || 3000));

// Body Parser Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
  extended: false
}));

// Get Static Path
app.use(express.static(path.join(__dirname, "public")));


app.post("/parse_excel", function (req, res) {

  upload(req, res, (err) => {
    if (err) {
      res.send({
        error: true,
        msg: err
      });
    } else {
      if (XLSX.read("./uploads/template.xlsx", {
          type: 'file'
        }).SheetNames.length > 1) {
        res.send({
          error: true,
          msg: "Please do not add any excel sheets to the template."
        });
      } else if (XLSX.read("./uploads/template.xlsx", {
          type: 'file'
        }).SheetNames[0] != "Sheet1") {
        res.send({
          error: true,
          msg: "Please do not change the name of the excel sheet"
        });
      }
      var workbook = XLSX.read("./uploads/template.xlsx", {
        type: 'file'
      });
      var first_sheet_name = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[first_sheet_name];
      var sendJSON = XLSX.utils.sheet_to_json(worksheet, {
        header: "A"
      })
      var keys = Object.keys(sendJSON[0]);
      var sample_names = [];
      var input_atttributes = {};
      for (var i = 0; i < keys.length; i++) {
        sample_names.push(sendJSON[0][keys[i]]);
        input_atttributes[sendJSON[0][keys[i]]] = [];
        for (var k = 1; k < sendJSON.length; k++) {
          input_atttributes[sendJSON[0][keys[i]]].push(sendJSON[k][keys[i]])
        }
      }
      fs.unlink("./uploads/template.xlsx", () => {
        res.send({
          error: false,
          data: {
            "sample_names": sample_names,
            "input_atttributes": input_atttributes
          }
        })
      })
    }
  });
});

app.listen(app.get("port"), function () {
  console.log("App is listenning o port 3000");
});